'***********************************************************************
'CREADO POR: 
'USUARIO CREACION : 
'FECHA CREACION:

'USUARIO MODIFICACION:CCORDOVA
'FECHA DE MODIFICACION:14/09/2012
'COMENTARIO MODIFICACION:
'MANEJO DE CANJE DE PROMOCIONES DE MARKETING

'***********************************************************************

Imports System.Data
Imports System.Data.SqlClient

Public Class clsCredito

    Dim oCons As ComConstante.clsConstante
    Dim fPersona As ComTipos.clsCargaPersona

    Public Function Cred_RecuperaValidacionesSimulador() As ComTipos.clsDataset


        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla("", "", "DevuelveValidacionesSimulador 10")

        Cred_RecuperaValidacionesSimulador = C

        C = Nothing
    End Function

    Public Function ValidaBaseDatosPersonasNegativas(ByVal pcDocumento As String) As ComTipos.clsDataset

        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla("", "", "Cred_ValidaBaseDatosPersonasNegativas '" & pcDocumento & "'")

        ValidaBaseDatosPersonasNegativas = C

        C = Nothing
    End Function
    Public Function Cred_RecuperaCreditosVigentes(ByVal pnCodPers As Integer, ByVal pnCodAge As Integer, ByVal pnMoneda As Integer) As ComTipos.clsDataset


        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla("", "", "Cred_RecuperaCreditosVigentes " & pnCodPers & "," & pnCodAge & "," & pnMoneda)

        Cred_RecuperaCreditosVigentes = C

        C = Nothing


    End Function

    Public Sub ActulizaTelefonosPersonas(ByVal pnCodPers As Integer, ByVal pcTelefono As String, ByVal pcMovil As String, ByVal pcTelefNego As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("ActualizaTelefonosPersonas ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar, 20).Value = pcTelefono
        oComm.Parameters.Add("@cMovil", SqlDbType.VarChar, 20).Value = pcMovil
        oComm.Parameters.Add("@cTelefNegocio", SqlDbType.VarChar, 20).Value = pcTelefNego
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub Graba_LLamadasCallCenter(ByVal pnCodPers As Integer, ByVal pnCodResult As Integer, _
        ByVal pcComentario As String, ByVal pdFecCompromiso As Date, ByVal pcTelefono As String, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pcContacto As String, ByVal nMontoCP As String, ByRef ngraba As Integer, ByVal psCodUsu As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        'Mestrada 01/09/2008 Agregue el Monto CP 
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("GrabaCallCenter ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nCodResult", SqlDbType.Int).Value = pnCodResult
        oComm.Parameters.Add("@cComentario", SqlDbType.NVarChar, 500).Value = pcComentario
        oComm.Parameters.Add("@dFechaCompromiso", SqlDbType.DateTime).Value = pdFecCompromiso
        oComm.Parameters.Add("@cNumTelef", SqlDbType.NVarChar, 50).Value = pcTelefono
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@cContacto", SqlDbType.VarChar, 250).Value = pcContacto
        oComm.Parameters.Add("@nMontoCP", SqlDbType.Money).Value = CDbl(nMontoCP)
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing
    End Sub


    Public Sub Graba_ImpresionDoc(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, _
          ByVal pcCodUsu As String, ByVal pnTipo As Integer, ByVal dFecDesemb As Date, ByVal pnCodGrp As Integer, ByVal pnCiclo As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Graba_ImpresionDoc ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 10).Value = pcCodUsu
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = pnTipo
        oComm.Parameters.Add("@dFecDesemb", SqlDbType.DateTime).Value = dFecDesemb
        oComm.Parameters.Add("@nCodGrp", SqlDbType.Int).Value = pnCodGrp
        oComm.Parameters.Add("@nCiclo", SqlDbType.Int).Value = pnCiclo
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing        
    End Sub

    Public Function ActualizaLineaCreditoTasas(ByVal pnTipo As Integer, ByVal pnCodLinea As Integer, _
            ByVal pnCodAutom As Double, ByVal pnTasaMin As Double, ByVal pnTasaMax As Double) As Integer


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizaDatosLineaCredTasas ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnTipo", SqlDbType.Int).Value = pnTipo
        oComm.Parameters.Add("@pnCodLinea", SqlDbType.Int).Value = pnCodLinea
        oComm.Parameters.Add("@pnCodAutonom", SqlDbType.Int).Value = pnCodAutom
        oComm.Parameters.Add("@pnTasaMin", SqlDbType.Money).Value = pnTasaMin
        oComm.Parameters.Add("@pnTasaMax", SqlDbType.Money).Value = pnTasaMax
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        ActualizaLineaCreditoTasas = oComm.Parameters(5).Value

        oComm = Nothing
        objCon = Nothing

    End Function

    Public Function RecuperaDatosPantallaLineaCredTasas(ByVal pnCodLinea As Integer) As ComTipos.ClsCredTarifarioTasa
        Dim C As ComTipos.ClsCredTarifarioTasa
        C = New ComTipos.ClsCredTarifarioTasa
        C.oCboAutonomias.CargaTabla("", "", "CredW_RecuperaCatalogoCodigos '4045'")
        C.oGridDatos.CargaTabla("", "", "Cred_RecuperaDatosLineaCredTasas " & pnCodLinea)
        RecuperaDatosPantallaLineaCredTasas = C

        C = Nothing

    End Function

    Public Function RegistraAnalistaAGestor(ByVal pnCodAge As Integer, ByVal pnCodpersGestor As Integer, ByVal pnCodPersAnalista As Integer) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraAnalistaAGestor ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodPersGestor", SqlDbType.Int).Value = pnCodpersGestor
        oComm.Parameters.Add("@pnCodPersAnalista", SqlDbType.Int).Value = pnCodPersAnalista
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        RegistraAnalistaAGestor = oComm.Parameters(3).Value

        oComm = Nothing
        objCon = Nothing

    End Function

    Public Sub RegistraCredparaAutorizar(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pcCodUsu As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EnviaCreditosParaAutorizar ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Sub RegistraCredGrupalparaAutorizar(ByVal pnCodGrupo As Integer, ByVal pnCiclo As Integer, ByVal pcCodUsu As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EnviaCreditosGrupalesParaAutorizar ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGrupo", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@nCiclo", SqlDbType.Int).Value = pnCiclo
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub


    Public Sub RegistraObsBandeja(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pcObs As String, ByVal pcCodUsu As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraObservaciones ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@cObs", SqlDbType.VarChar, 350).Value = pcObs
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function RecuperaDatosDeGestor(ByVal pnCodAge As Integer) As ComTipos.ClsParamGestor
        Dim C As ComTipos.ClsParamGestor
        C = New ComTipos.ClsParamGestor

        C.GrdAnalistasSINGestor.CargaTabla(, , " Cred_RecuperaAnalistaSinGestor " & pnCodAge)
        C.GrdAnalistasSINGestor.CargaTabla(, , " Cred_RecuperaAnalistaCONGestor " & pnCodAge)

        RecuperaDatosDeGestor = C

        C = Nothing
    End Function

    Public Function RegsitraGestorDEAnalista(ByVal pnCodAge As Integer, ByVal pnCodPersGestor As Integer, ByVal pnCodPersAnalista As Integer) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraAnalistaAGestor ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodPersGestor", SqlDbType.Int).Value = pnCodPersGestor
        oComm.Parameters.Add("@pnCodPersAnalista", SqlDbType.Int).Value = pnCodPersAnalista

        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        RegsitraGestorDEAnalista = oComm.Parameters(3).Value

        oComm = Nothing
        objCon = Nothing

    End Function


    Public Sub ActualizaValidacionSimulador(ByVal pnCodVar As Integer, ByVal psValorVar As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("ActualizarValidacionSimulador ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodVar", SqlDbType.Int).Value = pnCodVar
        oComm.Parameters.Add("@cValorVar", SqlDbType.Int).Value = psValorVar

        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Sub EliminaGestorDEAnalista(ByVal pnCodAge As Integer, ByVal pnCodPersGestor As Integer, ByVal pnCodPersAnalista As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EliminaAnalistaAGestor ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodPersGestor", SqlDbType.Int).Value = pnCodPersGestor
        oComm.Parameters.Add("@pnCodPersAnalista", SqlDbType.Int).Value = pnCodPersAnalista


        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Sub ValidaBaseDatosPEPs(ByVal pcNroDoc As String, ByVal pnModulo As Integer, ByVal pcUser As String, ByVal pnCodAge As Integer, ByRef pnValida As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ValidaBaseDatosPEPs ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cNroDoc", SqlDbType.VarChar, 13).Value = pcNroDoc
        oComm.Parameters.Add("@nModulo", SqlDbType.Int).Value = pnModulo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = pcUser
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nValida", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        pnValida = oComm.Parameters("@nValida").Value
        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Sub ValidaZonasAltoRiesgo(ByVal pcNroDoc As String, ByVal pnCodAge As Integer, ByVal pcUser As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ValidaZonasAltoRiesgo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cNroDoc", SqlDbType.VarChar, 12).Value = pcNroDoc
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@cCodusu", SqlDbType.VarChar, 20).Value = pcUser

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Sub ActualizaDataParaGestorDERecuperaciones(ByVal pnCodAge As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("ActualizaCarteraParaGestor ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge

        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaDatosParaFormularioGestor(ByVal pnCodAge As Integer, ByVal pnCodPersGestor As Integer) As ComTipos.ClsParamGestor
        Dim C As ComTipos.ClsParamGestor
        C = New ComTipos.ClsParamGestor

        C.GrdAnalistasSINGestor.CargaTabla(, , " Cred_RecuperaAnalistaSinGestor " & pnCodAge)
        C.GrdAnalistasConGestor.CargaTabla(, , " Cred_RecuperaAnalistaCONGestor " & pnCodAge & "," & pnCodPersGestor)
        'C.cboGestores.CargaTabla(, , " CredW_ListarGestores_x_Agencia " & pnCodAge)


        RecuperaDatosParaFormularioGestor = C

        C = Nothing
    End Function
    Public Sub ReprogramarENLote()
        Dim R As ComTipos.clsDataset
        Dim D As ComTipos.ClsCredReprog
        Dim DReprog As ComTipos.ClsCredReprog
        Dim C As ComTipos.ClsCalendario

        Dim i As Integer
        Dim nNumDias As Integer
        Dim dFecUno As Date
        Dim dFecDos As Date
        Dim nInteres As Double


        R = New ComTipos.clsDataset

        R.CargaTabla(, , "Cred_RecuperaCreditosParaReprogLote")
        R.Inicio()
        Do While Not R.EOF

            'Carga Datos
            D = RecuperaDatosReprogramacion(CInt(R.ObtenerDato("nCodAge")), CInt(R.ObtenerDato("nCodCred")))

            'Generar Nuevo Calendario
            If D.Calen.nNumFilas > 1 Then

                DReprog = New ComTipos.ClsCredReprog
                DReprog.Calen = D.Calen.DevolverCopia

                dFecUno = CDate(D.Calen.ObtenerDato("dFecVenc", 1))
                dFecDos = CDate(D.Calen.ObtenerDato("dFecVenc", 2))

                nNumDias = DateDiff(DateInterval.Day, dFecUno, dFecDos)
                C = New ComTipos.ClsCalendario
                nInteres = C.CalculaInteresComp(CDbl(D.Datos.ObtenerDato("nTasaComp")), nNumDias, CDbl(D.Datos.ObtenerDato("nSaldoK")))

                C = Nothing

                DReprog.Calen.Inicio()
                For i = 1 To DReprog.Calen.nNumFilas - 1
                    DReprog.Calen.AsignarDato(DReprog.Calen.ObtenerDato("dFecVenc", i + 1), "dFecVenc")
                    DReprog.Calen.Siguiente()
                Next
                i = DReprog.Calen.nNumFilas - 1
                DReprog.Calen.Seek(i)
                DReprog.Calen.AsignarDato(Trim(Str(CDbl(DReprog.Calen.ObtenerDato("nInteres")) + nInteres + CDbl(DReprog.Calen.ObtenerDato("nInteres", i + 1)))), "nInteres")
                DReprog.Calen.AsignarDato(Trim(Str(CDbl(DReprog.Calen.ObtenerDato("nCapital")) + CDbl(DReprog.Calen.ObtenerDato("nCapital", i + 1)))), "nCapital")
                DReprog.Calen.AsignarDato(Trim(Str(CDbl(DReprog.Calen.ObtenerDato("nIntGra")) + CDbl(DReprog.Calen.ObtenerDato("nIntGra", i + 1)))), "nIntGra")
                DReprog.Calen.AsignarDato(Trim(Str(CDbl(DReprog.Calen.ObtenerDato("nIntMor")) + CDbl(DReprog.Calen.ObtenerDato("nIntMor", i + 1)))), "nIntMor")
                DReprog.Calen.AsignarDato(Trim(Str(CDbl(DReprog.Calen.ObtenerDato("nIgv")) + CDbl(DReprog.Calen.ObtenerDato("nIgv", i + 1)))), "nIgv")
                DReprog.Calen.AsignarDato(Trim(Str(CDbl(DReprog.Calen.ObtenerDato("nGasto")) + CDbl(DReprog.Calen.ObtenerDato("nGasto", i + 1)))), "nGasto")

                DReprog.Calen.AsignarDato(Trim(Str(CDbl(DReprog.Calen.ObtenerDato("nCapital")) + CDbl(DReprog.Calen.ObtenerDato("nInteres")) + CDbl(DReprog.Calen.ObtenerDato("nIntGra")) + CDbl(DReprog.Calen.ObtenerDato("nIntMor")) + CDbl(DReprog.Calen.ObtenerDato("nIgv")) + CDbl(DReprog.Calen.ObtenerDato("nGasto")))), "nTotal")
                DReprog.Calen.EliminaRegistro(i + 1)


                RegistraReprogramacion(CInt(R.ObtenerDato("nCodAge")), CInt(R.ObtenerDato("nCodCred")), DReprog)
                DReprog = Nothing
            End If


            R.Siguiente()
        Loop

        R = Nothing


    End Sub
    Public Function RecuperaSubProductos(ByVal pnCodProd As Integer)
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaSubProductos " & pnCodProd)
        RecuperaSubProductos = D
        D = Nothing
    End Function

    Public Sub EliminaRolDEPersona(ByVal pnCodAge As Integer, ByVal pnCodPers As Integer, ByVal pnCodRol As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EliminaRol ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pnCodRol", SqlDbType.Int).Value = pnCodRol

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Sub RegistraRolDEPersona(ByVal pnCodAge As Integer, ByVal pnCodPers As Integer, ByVal pnCodRol As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraRol ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pnCodRol", SqlDbType.Int).Value = pnCodRol

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function RecuperaPersonasDERol(ByVal pnCodRol As Integer, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaPersonasDERol " & pnCodRol & "," & pnCodAge)
        RecuperaPersonasDERol = D
        D = Nothing


    End Function
    Public Sub ActualizaAgenciasBN(ByVal pnID As Long, ByVal pcCodDep As String, ByVal pcProvincia As String, _
    ByVal pcDistrito As String, ByVal pcNomAge As String, ByVal pcDireccion As String, ByVal pcTelefono As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizaAgenciasBN ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnID", SqlDbType.Int).Value = pnID
        oComm.Parameters.Add("@pcCodDep", SqlDbType.VarChar, 50).Value = pcCodDep
        oComm.Parameters.Add("@pcProvincia", SqlDbType.VarChar, 50).Value = pcProvincia
        oComm.Parameters.Add("@pcDistrito", SqlDbType.VarChar, 50).Value = pcDistrito
        oComm.Parameters.Add("@pcNomAge", SqlDbType.VarChar, 50).Value = pcNomAge
        oComm.Parameters.Add("@pcDireccion", SqlDbType.VarChar, 50).Value = pcDireccion
        oComm.Parameters.Add("@pcTelefono", SqlDbType.VarChar, 50).Value = pcTelefono


        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function RecuperaOficinasBN(ByVal psCodDep As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDepartamentos '" & psCodDep & "'")
        RecuperaOficinasBN = D
        D = Nothing
    End Function
    Public Function RecuperaDepartamentos() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDepartamentos ")
        RecuperaDepartamentos = D
        D = Nothing
    End Function

    Public Sub RegistraPlanInversion(ByVal pnCodEval As Integer, ByVal oData As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraPlanInversion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@pDatosGrilla", SqlDbType.Xml).Value = oData.DevuelveXML

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaPlanInversion(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredRecuperaPlanInv " & pnCodEval)
        RecuperaPlanInversion = D
        D = Nothing
    End Function


    Public Function ExaminarGrupos(ByVal pnTipoBusq As Integer, ByVal pnBusqueda As String, ByVal pnCodAge As Integer)
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_BuscaGrupos " & pnTipoBusq & ",'" & pnBusqueda & "'," & pnCodAge)
        ExaminarGrupos = D
        D = Nothing

    End Function

    Public Function RecuperaDatosCalendario(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDatosCalendario " & pnCodAge & "," & pnCodCred)
        RecuperaDatosCalendario = D
        D = Nothing
    End Function

    Public Function RecuperaDatosDesembolsosParciales(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))
        RecuperaDatosDesembolsosParciales = D
        D = Nothing
    End Function

    Public Function RecuperaCuotasAnteriores(ByVal pnCodAge As Integer, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_CuotasAnterioresxFechas " & pnCodAge & ",'" & pdFecIni & "', '" & pdFecFin & "'")
        RecuperaCuotasAnteriores = D
        D = Nothing
    End Function

    Public Function RecuperaCreditosCierreMes(ByVal pnCodAge As Integer, ByVal pnMes As Integer, ByVal pnAño As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_DevuelveCreditosCierreMes " & pnCodAge & "," & pnMes & ", " & pnAño)
        RecuperaCreditosCierreMes = D
        D = Nothing
    End Function

    Public Function RecuperaCreditosPlanillaPagos(ByVal pnCodAge As Integer, ByVal pdFecIni As Date, ByVal pdFecFin As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_CreditosPlanillaPagos " & pnCodAge & ",'" & pdFecIni & "', '" & pdFecFin & "'")
        RecuperaCreditosPlanillaPagos = D
        D = Nothing
    End Function

    Public Function RecuperaDatosCalendario_EvalConsumo(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDatosCalendarioEvalConsumo " & pnCodAge & "," & pnCodCred)
        RecuperaDatosCalendario_EvalConsumo = D
        D = Nothing
    End Function
    Public Function RecuperaConyugeCredInd(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaConyugeCredIndiv " & pnCodPers)
        RecuperaConyugeCredInd = D
        D = Nothing

    End Function
    Public Function RecuperaLineasCredito(ByVal pnCodAge As Integer, ByVal pnCodCamp As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaLineasDeCredito " & pnCodAge & "," & pnCodCamp)
        RecuperaLineasCredito = D
        D = Nothing
    End Function
    Public Function EliminaLineaCredito(ByVal nCodLinea As Integer, ByVal pnCodAge As Integer, _
        ByVal pnCodCamp As Integer) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EliminaLineaCredito ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodLinea", SqlDbType.Int).Value = nCodLinea
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCamp", SqlDbType.Int).Value = pnCodCamp
        oComm.Parameters.Add("@Resp", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        EliminaLineaCredito = oComm.Parameters(3).Value

        oComm = Nothing
        objCon = Nothing
    End Function

    Public Function ActualizaLineaCredito(ByVal nCodLinea As Integer, ByVal cDescLinea As String, _
        ByVal nTasaCom As Double, ByVal nTasaMor As Double, ByVal nPlazoMax As Integer, ByVal nPlazoMin As Integer, _
        ByVal nMontoMax As Double, ByVal nMontoMin As Double, ByVal nNumPresMin As Integer, ByVal nNumPresMax As Integer, _
        ByVal nMoneda As Integer, ByVal nProd As Integer, ByVal nSubProd As Integer, ByVal bEstado As Integer, ByVal pnCodAge As Integer, _
        ByVal pnCodCamp As Integer) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizaLineaCredito ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodLinea", SqlDbType.Int).Value = nCodLinea
        oComm.Parameters.Add("@cDescLinea", SqlDbType.VarChar, 100).Value = cDescLinea
        oComm.Parameters.Add("@nTasaCom", SqlDbType.Money).Value = nTasaCom
        oComm.Parameters.Add("@nTasaMor", SqlDbType.Money).Value = nTasaMor
        oComm.Parameters.Add("@nPlazoMax", SqlDbType.Int).Value = nPlazoMax
        oComm.Parameters.Add("@nPlazoMin", SqlDbType.Int).Value = nPlazoMin
        oComm.Parameters.Add("@nMontoMax", SqlDbType.Money).Value = nMontoMax
        oComm.Parameters.Add("@nMontoMin", SqlDbType.Money).Value = nMontoMin
        oComm.Parameters.Add("@nNumPresMin", SqlDbType.Int).Value = nNumPresMin
        oComm.Parameters.Add("@nNumPresMax", SqlDbType.Int).Value = nNumPresMax
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = nProd
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = nSubProd
        oComm.Parameters.Add("@bEstado", SqlDbType.Int).Value = bEstado
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCamp", SqlDbType.Int).Value = pnCodCamp
        oComm.Parameters.Add("@Resp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()
        ActualizaLineaCredito = oComm.Parameters(16).Value

        oComm = Nothing
        objCon = Nothing
    End Function
    Public Function RecuperaControlesTarifario(ByVal pdFecsist As Date) As ComTipos.ClsCredTarifario
        Dim C As ComTipos.ClsCredTarifario
        C = New ComTipos.ClsCredTarifario
        C.cboAgencia.CargaTabla(, , "CredW_Recupera_Agencias ")
        C.cboCampania.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecsist, "dd/MM/yyyy") & "'")
        C.cboMoneda.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '2090' ")
        C.cboProducto.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '1500' ")
        C.cboSubProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '1600' ")

        RecuperaControlesTarifario = C
        C = Nothing
    End Function

    Public Function RecuperaCicloDEGrupo(ByVal pnCodAge As Integer, ByVal pnCodGrupo As Integer) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RecuperaCiclodeGrupo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodGrp", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@pnCicloResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        RecuperaCicloDEGrupo = oComm.Parameters(2).Value

        oComm = Nothing
        objCon = Nothing
    End Function

    Public Sub ActualizaDatosGrupo(ByVal pnCodAge As Integer, ByVal pnCodGrupo As Integer, ByVal pcNomGrupo As String, _
            ByVal pnCiclo As Integer, ByVal pnTipoGrupo As Integer, ByVal pcDireccReunion As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegsitraDatosGrupo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodGrupo", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@pcNomGrupo", SqlDbType.VarChar, 150).Value = pcNomGrupo
        oComm.Parameters.Add("@pnCiclo", SqlDbType.Int).Value = pnCiclo
        oComm.Parameters.Add("@pnTipoGrupo", SqlDbType.Int).Value = pnTipoGrupo
        oComm.Parameters.Add("@pcDireccReunion", SqlDbType.VarChar, 1000).Value = pcDireccReunion


        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing
    End Sub


    Public Sub EliminaIntegranteGrupo(ByVal pnCodAge As Integer, ByVal pnCodGrupo As Integer, ByVal pnCodPers As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EliminaIntegranteGrupo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodGrupo", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing
    End Sub
    Public Sub ActualizaIntegrante(ByVal pnCodAge As Integer, ByVal pnCodGrupo As Integer, ByVal pnCodPers As Integer, ByVal pnTiporela As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizaIntegranteGrupo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodGrupo", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pnTiporela", SqlDbType.Int).Value = pnTiporela
        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub RegistraIntegrante(ByVal pnCodAge As Integer, ByVal pnCodGrupo As Integer, ByVal pnCodPers As Integer, ByVal pnTiporela As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraIntegranteGrupo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodGrupo", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pnTiporela", SqlDbType.Int).Value = pnTiporela
        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing
    End Sub
    Public Function RecuperaIntegrantesDeGrupo(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_RecuperaIntegrantesDeGrupo " & pnCodAge & "," & pnCodCred)
        RecuperaIntegrantesDeGrupo = C
        C = Nothing
    End Function

    Public Function RecuperaDatosGrupo(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsCredGrupo
        Dim C As ComTipos.ClsCredGrupo
        C = New ComTipos.ClsCredGrupo
        C.DatosGrupo.CargaTabla(, , "Cred_RecuperaDatosDeGrupo " & pnCodAge & "," & pnCodCred)
        C.DatosInteg.CargaTabla(, , "Cred_RecuperaIntegrantesDeGrupo " & pnCodAge & "," & pnCodCred)
        RecuperaDatosGrupo = C
        C = Nothing
    End Function

    Public Function RecuperaCiclosGrupo(ByVal pnCodAge As Integer, _
                                            ByVal pnCodGrupo As Integer, _
                                            ByVal pnEstado As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_RecuperaCiclosGrupo " & pnCodAge & "," & pnCodGrupo & "," & pnEstado)
        RecuperaCiclosGrupo = C
        C = Nothing
    End Function

    Public Function RecuperaControlesMantGrupos() As ComTipos.ClsCredGrupo
        Dim C As ComTipos.ClsCredGrupo
        C = New ComTipos.ClsCredGrupo
        C.cboTipoGrupo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '2170'")
        C.cboCargosEnGrupo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '2171'")
        RecuperaControlesMantGrupos = C
        C = Nothing

    End Function

    Public Function RecuperaControlsVisita() As ComTipos.ClsCredVisita
        Dim D As ComTipos.ClsCredVisita
        D = New ComTipos.ClsCredVisita

        D.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '2167'")
        D.CboConVisita.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '2168'")
        D.CboAteCli.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '2169'")

        RecuperaControlsVisita = D

    End Function

    Public Sub RegistrarVisita(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal poDatos As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraVisitaCred ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@dFecVisita", SqlDbType.DateTime).Value = CDate(poDatos.ObtenerDato("dFecVisita"))
        oComm.Parameters.Add("@nDestCred", SqlDbType.Int).Value = CInt(poDatos.ObtenerDato("nDestCred"))
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 5000).Value = poDatos.ObtenerDato("cComentario")
        oComm.Parameters.Add("@nConclVisita", SqlDbType.Int).Value = CInt(poDatos.ObtenerDato("nConclVisita"))
        oComm.Parameters.Add("@nConclAte", SqlDbType.Int).Value = CInt(poDatos.ObtenerDato("nConclAte"))
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = CInt(poDatos.ObtenerDato("nCodPersVisita"))

        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Sub EliminarVisita(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EliminaVisitaCred ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.ExecuteNonQuery()


        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaVisita(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaVisitaCred " & pnCodAge & "," & pnCodCred)
        RecuperaVisita = D
        D = Nothing
    End Function
    Public Function ObtieneDeudaAFecha(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pdFecha As Date) As Double
        Dim R As ComTipos.clsDataset

        R = New ComTipos.clsDataset
        R.CargaTabla(, , "Cred_ConsCredDeudaFecha '" & Trim(Str(pnCodAge)) & "','" & Trim(Str(pnCodCred)) & "','" & pdFecha & "'")

        ObtieneDeudaAFecha = CDbl(R.ObtenerDato("nMontoDeudaFecha"))

        R = Nothing
    End Function

    Public Function ObtieneGastoTotalCal(ByVal pnPrestamo As Double, ByVal pnMoneda As Integer, _
        ByVal pnProd As Integer, ByVal pnSubProd As Integer, Optional ByVal pnSegInd As Integer = 1) As Double

        Dim R As ComTipos.clsDataset

        R = New ComTipos.clsDataset
        R.CargaTabla(, , "CredW_RecuperaGastosParaSimulador " & pnPrestamo & "," & pnMoneda & "," & pnProd & "," & pnSubProd & "," & pnSegInd)

        ObtieneGastoTotalCal = CDbl(R.ObtenerDato("nValorGasto"))

        R = Nothing

    End Function

    Public Function RecuperaTasaSeguroDesgravamenSimCal(ByVal pnMoneda As Integer, ByVal pnPrestamo As Double, _
        ByVal pnSubProd As Integer, ByVal pnProd As Integer, ByVal pnIndiv As Integer, ByVal pnCodAge As Integer) As Double

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RecuperaTasaSeguroDesgravamenSimCalen ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nValor", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nMonedaCred", SqlDbType.Int).Value = pnMoneda
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = pnPrestamo
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = pnSubProd
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = pnProd
        oComm.Parameters.Add("@nIndiv", SqlDbType.Int).Value = pnIndiv
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.ExecuteNonQuery()

        RecuperaTasaSeguroDesgravamenSimCal = oComm.Parameters(0).Value

        oComm = Nothing
        objCon = Nothing

    End Function

    Public Function RecuperaTasaSeguroDesgravamen(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, _
                                            ByVal pnPrestamo As Double, ByVal pnMoneda As Integer, _
                                                Optional ByVal pnProd As Integer = 0, Optional ByVal pnSubProd As Integer = 0) As Double

        Dim R As ComTipos.clsDataset

        R = New ComTipos.clsDataset
        R.CargaTabla(, , "Cred_RecuperaTasaSeguroDesgravamen " & pnCodAge & "," & pnCodCred & "," & pnPrestamo & "," & pnMoneda & "," & pnProd & "," & pnSubProd)

        RecuperaTasaSeguroDesgravamen = CDbl(R.ObtenerDato("nValor"))

        R = Nothing

    End Function

    Public Function RecuperaFeriados(ByVal pdFechaDesem As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_DevuelveFeriados '" & Format(pdFechaDesem, "dd/MM/yyyy") & "'")
        RecuperaFeriados = D
        D = Nothing
    End Function

    Public Function RecuperaPolizaMicroSeguro(ByVal pcSolicitud As String) As ComTipos.clsMicroSeguros
        Dim D As ComTipos.clsMicroSeguros
        D = New ComTipos.clsMicroSeguros
        D.Data.CargaTabla(, , "Cred_RecuperaMicroSeguros '" & pcSolicitud & "'")
        D.Beneficiarios.CargaTabla(, , "Cred_BeneficiariosMicroSeguros '" & pcSolicitud & "'")
        RecuperaPolizaMicroSeguro = D
        D = Nothing
    End Function
    Public Function RecuperaPersonasCallCenter(ByVal cCodUsu As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevuelveListaClientesCallCenter " & cCodUsu)
        RecuperaPersonasCallCenter = D
        D = Nothing
    End Function

    Public Function RecuperaResultadosLlamadas() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5105 ")
        RecuperaResultadosLlamadas = D
        D = Nothing
    End Function
    Public Function RecuperaCreditosAmortizados(ByVal pnCodAge As Integer, ByVal pdFecha As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_CreditosAmortizados " & pnCodAge & ", '" & Format(pdFecha, "dd/MM/yyyy") & "'")
        RecuperaCreditosAmortizados = D
        D = Nothing
    End Function

    Public Function RecuperaDuplicadoCalend(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsCredDupCal
        Dim D As ComTipos.ClsCredDupCal
        D = New ComTipos.ClsCredDupCal
        D.oDato.CargaTabla(, , "Cred_DuplicCalendDatos " & pnCodAge & "," & pnCodCred)
        D.oCalend.CargaTabla(, , "Cred_DuplicCalend " & pnCodAge & "," & pnCodCred)
        D.oCalenDesemb.CargaTabla(, , "Cred_RecuperaCalDesParcialDetalle " & pnCodAge & "," & pnCodCred)
        RecuperaDuplicadoCalend = D
        D = Nothing
    End Function

    Public Function RecuperaHojaResumen(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_HojaResumen " & pnCodAge & "," & pnCodCred)
        RecuperaHojaResumen = D
        D = Nothing
    End Function

    Public Function RecuperaCalendarioPendiente(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaCalendario " & pnCodAge & "," & pnCodCred)
        RecuperaCalendarioPendiente = D
        D = Nothing
    End Function

    Public Function RecuperaCalendarioCierreMes(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_CalendarioCuotasPagadas_V2 " & pnCodAge & "," & pnCodCred)
        RecuperaCalendarioCierreMes = D
        D = Nothing
    End Function

    Public Function RecuperaTasaAprobacionAsignada(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "RecuperaTasaAprobacionAsignada " & pnCodAge & "," & pnCodCred)
        RecuperaTasaAprobacionAsignada = D
        D = Nothing
    End Function
    Public Function RecuperaPersonasRelacionadas(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnRelacion As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_DevuelvePersonasRelacionadas " & pnCodAge & "," & pnCodCred & ", " & pnRelacion)
        RecuperaPersonasRelacionadas = D
        D = Nothing
    End Function

    Public Function ImpresionResumenComite(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnCodInst As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_ImpresionResolComite " & pnCodAge & "," & pnCodCred & "," & pnCodInst)
        ImpresionResumenComite = D
        D = Nothing

    End Function

    Public Function RecuperaExcedente(ByVal pnCodEval As Integer) As Double
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        RecuperaExcedente = 0

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RecuperaExcedente ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@pnExedente", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        RecuperaExcedente = oComm.Parameters("@pnExedente").Value

        oComm = Nothing
        objCon = Nothing

    End Function
    Public Function MontoDEGarantiaDelCredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pdFecSis As Date) As Double
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        MontoDEGarantiaDelCredito = 0

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_MontoDEGarantiaDelCredito ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pdFecSis", SqlDbType.DateTime).Value = pdFecSis
        oComm.Parameters.Add("@pnResp", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        MontoDEGarantiaDelCredito = oComm.Parameters("@pnResp").Value

        oComm = Nothing
        objCon = Nothing


    End Function


    Public Function CreditoTieneEvaluacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Boolean
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        CreditoTieneEvaluacion = False

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_TieneEvaluacion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        If oComm.Parameters("@pnResul").Value = 1 Then
            CreditoTieneEvaluacion = True
        Else
            CreditoTieneEvaluacion = False
        End If

        oComm = Nothing
        objCon = Nothing


    End Function

    Public Function RecuperaSolCredImp(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsImpCredSol
        Dim D As ComTipos.ClsImpCredSol
        D = New ComTipos.ClsImpCredSol
        D.oDatos.CargaTabla(, , "Imp_RecupSolCredDatos " & pnCodAge & "," & pnCodCred)
        D.oPersonas.CargaTabla(, , "Imp_RecupSolCredPersRel " & pnCodAge & "," & pnCodCred)
        D.oCalDesem.CargaTabla(, , "Cred_RecuperaCalendDesembolsos " & pnCodAge & "," & pnCodCred)
        RecuperaSolCredImp = D
        D = Nothing
    End Function

    Public Function RecuperaMotivosRechazoCredito() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '3002'")
        RecuperaMotivosRechazoCredito = D
        D = Nothing
    End Function

    Public Function RecuperaDatosRechazoCredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDatosBasicosDECredito  " & pnCodAge & "," & pnCodCred)
        RecuperaDatosRechazoCredito = D
        D = Nothing
    End Function

    Public Function CreditoAutorizado(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_CreditoAutorizado ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        CreditoAutorizado = oComm.Parameters("@pnResul").Value

        oComm = Nothing
        objCon = Nothing


    End Function

    Public Sub RegistraRechazoCredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnMotivo As Integer, ByVal psComen As String, ByVal pnObservar As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistroRechazoCredito ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnCodMotivo", SqlDbType.Int).Value = pnMotivo
        oComm.Parameters.Add("@psComen", SqlDbType.Text).Value = psComen
        oComm.Parameters.Add("@pnObservar", SqlDbType.Int).Value = pnObservar
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing


    End Sub

    Public Sub RegistraPerdonarMora(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pdDatos As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraPerdonarMora ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnDatos", SqlDbType.Xml).Value = pdDatos.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Sub RegistraMicroSeguros(ByVal pcPoliza As String, ByVal pcSolicitud As String, ByVal pnCodAge As Integer, ByVal pdFechaVig As Date, _
    ByVal pnCodPers As Integer, ByVal pnTipoPlan As Integer, ByVal pcCodUsu As String, ByVal pcNomCoasegurado As String, ByVal pnTipoDocCoasegurado As Integer, ByVal pcDniCoasegurado As String, _
    ByVal pdFechaNacCoasegurado As Date, ByVal pnEdadCoasegurado As Integer, ByVal pnSexoCoasegurado As Integer, ByVal pnParentCoasegurado As Integer, _
    ByVal pcTutor As String, ByVal pcDniTutor As String, ByVal oBeneficiarios As ComTipos.clsDataset, ByVal pnOpcion As Integer, ByRef pnValida As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_GrabaMicroSeguro ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nPoliza", SqlDbType.VarChar, 8).Value = pcPoliza
        oComm.Parameters.Add("@nSolicitud", SqlDbType.VarChar, 8).Value = pcSolicitud
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@dFechaVig", SqlDbType.DateTime).Value = pdFechaVig
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nTipoPlan", SqlDbType.Int).Value = pnTipoPlan
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@cNomCoasegurado", SqlDbType.VarChar, 80).Value = pcNomCoasegurado
        oComm.Parameters.Add("@nTipoDocCoasegurado", SqlDbType.Int).Value = pnTipoDocCoasegurado
        oComm.Parameters.Add("@cDniCoasegurado", SqlDbType.VarChar, 13).Value = pcDniCoasegurado
        oComm.Parameters.Add("@dFechaNacCoasegurado", SqlDbType.DateTime).Value = pdFechaNacCoasegurado
        oComm.Parameters.Add("@nEdadCoasegurado", SqlDbType.Int).Value = pnEdadCoasegurado
        oComm.Parameters.Add("@nSexoCoasegurado", SqlDbType.Int).Value = pnSexoCoasegurado
        oComm.Parameters.Add("@nParentCoasegurado", SqlDbType.Int).Value = pnParentCoasegurado
        oComm.Parameters.Add("@cTutor", SqlDbType.VarChar, 60).Value = pcTutor
        oComm.Parameters.Add("@cDniTutor", SqlDbType.VarChar, 13).Value = pcDniTutor
        oComm.Parameters.Add("@oBeneficiarios", SqlDbType.Xml).Value = oBeneficiarios.DevuelveXML
        oComm.Parameters.Add("@nOpcion", SqlDbType.Int).Value = pnOpcion
        oComm.Parameters.Add("@nValida", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        pnValida = oComm.Parameters.Item("@nValida").Value

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub RegistraReprogramacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pdRepro As ComTipos.ClsCredReprog)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraReprogramacion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@Calen", SqlDbType.Xml).Value = pdRepro.Calen.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaDatosReprogramacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsCredReprog
        Dim D As ComTipos.ClsCredReprog
        D = New ComTipos.ClsCredReprog

        D.Datos.CargaTabla(, , "Cred_RecuperaDatosBasicosDECredito  " & pnCodAge & "," & pnCodCred)
        D.Calen.CargaTabla(, , "Cred_RecuperaReprog_CuotasPend " & pnCodCred & "," & pnCodAge)

        RecuperaDatosReprogramacion = D
        D = Nothing

    End Function

    Public Function RecuperaConsultaCredDatos(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal psCartera As String) As ComTipos.ClsCreConsulta
        Dim D As ComTipos.ClsCreConsulta
        D = New ComTipos.ClsCreConsulta

        If psCartera = "FIDEICOMISO" Then
            D.Data.CargaTabla(, , "DBCORE..Cred_RecuperaDatosConsCred " & pnCodCred & "," & pnCodAge)
            D.CuotaPag.CargaTabla(, , "DBCORE..Cred_RecuperaConsCred_CuotasPag " & pnCodCred & "," & pnCodAge)
            D.CuotaPend.CargaTabla(, , "DBCORE..Cred_RecuperaConsCred_CuotasPend " & pnCodCred & "," & pnCodAge)
            D.DataDeudCal.CargaTabla(, , "DBCORE..Cred_RecuperaConsCred_DeudaTotalCal " & pnCodCred & "," & pnCodAge)
            D.Kardex.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredKardex " & pnCodCred & "," & pnCodAge)
            D.DataDeudaFecha.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredDeudaFecha " & pnCodCred & "," & pnCodAge)
            D.Garantias.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredGarantia " & pnCodCred & "," & pnCodAge)
            D.DataRelaPers.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredRelaPers " & pnCodCred & "," & pnCodAge)
            D.DatosCalendario.CargaTabla(, , "DBCORE..Cred_RecuperaDatosBasicosCalendario " & pnCodAge & "," & pnCodCred)
            D.CalendDesembolsos.CargaTabla(, , "DBCORE..Cred_RecuperaCalendDesembolsos " & pnCodAge & "," & pnCodCred)
        Else
            D.Data.CargaTabla(, , "DBSOLIDA..Cred_RecuperaDatosConsCred " & pnCodCred & "," & pnCodAge)
            D.CuotaPag.CargaTabla(, , "DBSOLIDA..Cred_RecuperaConsCred_CuotasPag " & pnCodCred & "," & pnCodAge)
            D.CuotaPend.CargaTabla(, , "DBSOLIDA..Cred_RecuperaConsCred_CuotasPend " & pnCodCred & "," & pnCodAge)
            D.DataDeudCal.CargaTabla(, , "DBSOLIDA..Cred_RecuperaConsCred_DeudaTotalCal " & pnCodCred & "," & pnCodAge)
            D.Kardex.CargaTabla(, , "DBSOLIDA..Cred_RecuperaConsCredKardex " & pnCodCred & "," & pnCodAge)
            D.DataDeudaFecha.CargaTabla(, , "DBSOLIDA..Cred_RecuperaConsCredDeudaFecha " & pnCodCred & "," & pnCodAge)
            D.Garantias.CargaTabla(, , "DBSOLIDA..Cred_RecuperaConsCredGarantia " & pnCodCred & "," & pnCodAge)
            D.DataRelaPers.CargaTabla(, , "DBSOLIDA..Cred_RecuperaConsCredRelaPers " & pnCodCred & "," & pnCodAge)
            D.DatosCalendario.CargaTabla(, , "DBSOLIDA..Cred_RecuperaDatosBasicosCalendario " & pnCodAge & "," & pnCodCred)
            D.CalendDesembolsos.CargaTabla(, , "DBSOLIDA..Cred_RecuperaCalendDesembolsos " & pnCodAge & "," & pnCodCred)
        End If
        RecuperaConsultaCredDatos = D
        D = Nothing
    End Function

    Public Function RegistraAutorizacionDesembParcial(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal psCodUsu As String, ByVal psComent As String, ByVal pnNumDesemb As Integer) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraAutorizacionDesParcial ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pcCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psComent", SqlDbType.VarChar, 5000).Value = psComent
        oComm.Parameters.Add("@pnNumDesem", SqlDbType.Int).Value = pnNumDesemb
        oComm.ExecuteNonQuery()

        RegistraAutorizacionDesembParcial = oComm.Parameters("@pnResul").Value

        oComm = Nothing
        objCon = Nothing

    End Function

    Public Function RegistraAutorizacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal psCodUsu As String, ByVal psComent As String) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraAutorizacion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pcCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psComent", SqlDbType.VarChar, 5000).Value = psComent
        oComm.ExecuteNonQuery()

        RegistraAutorizacion = oComm.Parameters("@pnResul").Value

        oComm = Nothing
        objCon = Nothing

    End Function

    Public Function RecuperaDatosAutorizacionDesPar(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsAutorizaCred
        Dim D As ComTipos.ClsAutorizaCred
        D = New ComTipos.ClsAutorizaCred
        D.oDatos = RecuperaDatosBasicosDECredito(pnCodAge, pnCodCred)
        D.oGridAutor.CargaTabla(, , "Cred_RecuperaDemParAutorizaciones " & pnCodCred & "," & pnCodAge)
        RecuperaDatosAutorizacionDesPar = D
        D = Nothing

    End Function

    Public Function RecuperaDatosAutorizacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsAutorizaCred
        Dim D As ComTipos.ClsAutorizaCred
        D = New ComTipos.ClsAutorizaCred
        D.oDatos = RecuperaDatosBasicosDECredito(pnCodAge, pnCodCred)
        D.oGridAutor.CargaTabla(, , "Cred_RecuperaAutorizaciones " & pnCodCred & "," & pnCodAge)
        RecuperaDatosAutorizacion = D
        D = Nothing

    End Function

    Public Function RecuperaEvaluacionesPersona(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_RecuperaEvaluaciones  " & pnCodPers)
        RecuperaEvaluacionesPersona = D
        D = Nothing

    End Function

    Public Function RecuperaDatosBasicosDECredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDatosBasicosDECredito  " & pnCodAge & "," & pnCodCred)
        RecuperaDatosBasicosDECredito = D
        D = Nothing

    End Function

    Public Sub RegistraNivelesDeRangos(ByVal pnCodRango As Integer, ByVal pData As ComTipos.clsDataset)
        '
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraNivelesDeRangos ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodRango", SqlDbType.Int)
        oComm.Parameters("@pnCodRango").Value = pnCodRango
        oComm.Parameters.Add("@pNiveles", SqlDbType.Xml)
        oComm.Parameters("@pNiveles").Value = pData.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaNivelesDERango(ByVal pnCodRango As Integer) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaNivelesDERangos " & pnCodRango)
        RecuperaNivelesDERango = D
        D = Nothing
    End Function
    Public Sub RegistraRangos(ByVal poData As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraRangos", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pRangos", SqlDbType.Xml)
        oComm.Parameters("@pRangos").Value = poData.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Sub RegistraNiveles(ByVal poData As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraNiveles", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pNiveles", SqlDbType.Xml)
        oComm.Parameters("@pNiveles").Value = poData.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub


    Public Sub RegistraUsuarioDeNiveles(ByVal pnNivel As Integer, ByVal poData As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraUsuariosDeNivel", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnNivel", SqlDbType.Int)
        oComm.Parameters("@pnNivel").Value = pnNivel
        oComm.Parameters.Add("@pDatos", SqlDbType.Xml)
        oComm.Parameters("@pDatos").Value = poData.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function RecuperaUsuariosDeNivel(ByVal pnCodNivel As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaUsuariosDeNivel " & pnCodNivel)
        RecuperaUsuariosDeNivel = D
        D = Nothing

    End Function

    Public Function RecuperaNiveles() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaNiveles")
        RecuperaNiveles = D
        D = Nothing

    End Function

    Public Function RecuperaNivelesYRangos() As ComTipos.ClsAutonomia
        Dim D As ComTipos.ClsAutonomia
        D = New ComTipos.ClsAutonomia
        D.oNiveles.CargaTabla(, , "Cred_RecuperaNiveles")
        D.oRangos.CargaTabla(, , "Cred_RecuperaRangos")
        D.oAgencias.CargaTabla(, , "CredW_Recupera_Agencias")
        D.oProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '1500'")
        D.oSubProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '1600'")
        D.oMoneda.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '2090'")

        RecuperaNivelesYRangos = D
        D = Nothing

    End Function
    Public Function RecuperaAsiento(ByVal pdFecha As Date, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim oDato As ComTipos.clsDataset
        oDato = New ComTipos.clsDataset
        oDato.CargaTabla(, , "Cred_RecuperaAsientoContable '" & Format(pdFecha, "dd/MM/yyyy") & "'," & pnCodAge, 3)
        RecuperaAsiento = oDato
        oDato = Nothing
    End Function

    Public Function RecuperaDatosSimulador(ByVal nPrestamo As Double, ByVal nMoneda As Integer, ByVal nProd As Integer, _
        ByVal nSubProd As Integer, ByVal nTercero As Integer, ByVal nPlazo As Integer, ByVal nIndiv As Integer) As ComTipos.clsDataset
        Dim oDato As ComTipos.clsDataset
        oDato = New ComTipos.clsDataset
        oDato.CargaTabla(, , "DevuelveDatosSimulador " & nPrestamo & "," & nMoneda & "," & nProd & ", " & nSubProd & "," & nTercero & "," & nPlazo & "," & nIndiv)
        RecuperaDatosSimulador = oDato
        oDato = Nothing
    End Function

    Public Function RecuperaAsientoFinMes(ByVal pdFecha As Date, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim oDato As ComTipos.clsDataset
        oDato = New ComTipos.clsDataset
        oDato.CargaTabla(, , "Cred_RecuperaAsientoFinMes '" & Format(pdFecha, "dd/MM/yyyy") & "'," & pnCodAge, 3)
        RecuperaAsientoFinMes = oDato
        oDato = Nothing
    End Function

    Public Function ValidaAsiento(ByVal pdFecha As Date) As Boolean
        Dim oDato As ComTipos.clsDataset
        oDato = New ComTipos.clsDataset
        oDato.CargaTabla(, , "Cred_VerficaAsientoOK '" & Format(pdFecha, "dd/MM/yyyy") & "'")
        If oDato.nNumFilas > 0 Then
            If oDato.ObtenerDato("bOK") = "1" Then
                ValidaAsiento = True
            Else
                ValidaAsiento = False
            End If
        Else
            ValidaAsiento = False
        End If
        oDato = Nothing
    End Function

    Public Sub CargarDatosFormPagoCuotas(ByRef P As ComTipos.clsCredPagoCuotas)
        oCons = New ComConstante.clsConstante
        P.Bancos = oCons.RecuperaConstantes(2000)
        oCons = Nothing
    End Sub

    Public Function CargaDatosConsltaGrupal(ByVal pnCodGrupo As Integer, ByVal pnCiclo As Integer) As ComTipos.ClsConsGrupal
        Dim D As ComTipos.ClsConsGrupal
        D = New ComTipos.ClsConsGrupal
        D.oIntegrantes.CargaTabla(, , "Cred_RecuperaConsultaGrpIntegrante " & pnCodGrupo & "," & pnCiclo)
        If D.oIntegrantes.nNumFilas > 0 Then
            D.oData.CargaTabla(, , "Cred_RecuperaDatosConsGrp " & D.oIntegrantes.ObtenerDato("nCodCred") & "," & D.oIntegrantes.ObtenerDato("nCodAge"))
        End If

        CargaDatosConsltaGrupal = D
    End Function

    Public Function CargaDatosAprobacionGrupalNew(ByVal pnCodAge As Integer, ByVal pdFecSis As Date, ByVal pnCodGrp As Integer, ByVal pnCiclo As Integer, ByVal pnAutonom As Integer) As ComTipos.ClsWCredSolGrp
        Dim D As ComTipos.ClsWCredSolGrp
        D = New ComTipos.ClsWCredSolGrp

        'D = NuevoCreditoGrupal(pnCodAge, pdFecSis)
        D.Integrantes.CargaTabla(, , "Cred_RecuperaSolGrpIntegrantesAprobacionNew " & pnCodGrp & "," & pnCiclo & "," & pnAutonom)
        If D.Integrantes.nNumFilas > 0 Then
            D.Integrantes.Inicio()
            Do While Not D.Integrantes.EOF
                D.Integrantes.AsignarDato(D.SeparaActividad(D.Integrantes.ObtenerDato("cActividad")), "cActividad")
                D.Integrantes.Siguiente()
            Loop
            D.Integrantes.Seek(1)
            D.Datos.CargaTabla(, , "CredW_RecuperaDatosAprGrp " & D.Integrantes.ObtenerDato("nCodCred") & "," & pnCodAge)
        End If
        CargaDatosAprobacionGrupalNew = D

    End Function

    Public Function CargaDatosAprobacionGrupal(ByVal pnCodAge As Integer, ByVal pdFecSis As Date, ByVal pnCodGrp As Integer, ByVal pnCiclo As Integer, ByVal pnAutonom As Integer) As ComTipos.ClsWCredSolGrp
        Dim D As ComTipos.ClsWCredSolGrp
        D = New ComTipos.ClsWCredSolGrp

        'D = NuevoCreditoGrupal(pnCodAge, pdFecSis)
        D.Integrantes.CargaTabla(, , "Cred_RecuperaSolGrpIntegrantesAprobacion " & pnCodGrp & "," & pnCiclo & "," & pnAutonom)
        If D.Integrantes.nNumFilas > 0 Then
            D.Integrantes.Inicio()
            Do While Not D.Integrantes.EOF
                D.Integrantes.AsignarDato(D.SeparaActividad(D.Integrantes.ObtenerDato("cActividad")), "cActividad")
                D.Integrantes.Siguiente()
            Loop
            D.Integrantes.Seek(1)
            D.Datos.CargaTabla(, , "CredW_RecuperaDatosAprGrp " & D.Integrantes.ObtenerDato("nCodCred") & "," & pnCodAge)
        End If
        CargaDatosAprobacionGrupal = D

    End Function

    Public Function CargaDatosSolicitudGrupal(ByVal pnCodAge As Integer, ByVal pdFecSis As Date, ByVal pnCodGrp As Integer, ByVal pnCiclo As Integer) As ComTipos.ClsWCredSolGrp
        Dim D As ComTipos.ClsWCredSolGrp
        D = New ComTipos.ClsWCredSolGrp

        'D = NuevoCreditoGrupal(pnCodAge, pdFecSis)
        D.Integrantes.CargaTabla(, , "CredW_RecuperaSolGrpIntegrantes " & pnCodGrp & "," & pnCiclo)
        If D.Integrantes.nNumFilas > 0 Then
            D.Integrantes.Inicio()
            Do While Not D.Integrantes.EOF
                D.Integrantes.AsignarDato(D.SeparaActividad(D.Integrantes.ObtenerDato("cActividad")), "cActividad")
                D.Integrantes.Siguiente()
            Loop
            D.Integrantes.Seek(1)
            D.Datos.CargaTabla(, , "CredW_RecuperaDatosSolGrp " & D.Integrantes.ObtenerDato("nCodCred") & "," & pnCodAge)
        End If
        CargaDatosSolicitudGrupal = D
    End Function

    Public Function CargaGrupos(ByVal pnCodAge As Integer) As ComTipos.ClsWCredSolGrp
        Dim D As ComTipos.ClsWCredSolGrp
        D = New ComTipos.ClsWCredSolGrp
        D.Grupo.CargaTabla(, , "CredW_RecuperaGrupos " & Trim(Str(pnCodAge)))
        CargaGrupos = D
        D = Nothing
    End Function



    Public Function NuevoCreditoGrupal(ByVal pnCodAge As Integer, ByVal pdFecSis As Date) As ComTipos.ClsWCredSolGrp
        Dim D As ComTipos.ClsWCredSolGrp
        D = New ComTipos.ClsWCredSolGrp
        D.Analistas.CargaTabla(, , "CredW_RecuperaAnalista  " & Trim(Str(pnCodAge)))
        D.Oficinas.CargaTabla(, , "Gen_RecuperaOficinasPorAgencia  " & Trim(Str(pnCodAge)))
        D.Producto.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        D.SubProducto.CargaTabla(, , "Cred_Recupera_SubProdGrupal 1")
        D.Condicion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2070")
        ' D.Grupo.CargaTabla(, , "CredW_RecuperaGrupos " & Trim(Str(pnCodAge)))
        D.Campania.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        D.Destinos.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        D.Actividades.CargaTabla(, , "CredW_RecuperaActividades")
        NuevoCreditoGrupal = D
        D = Nothing

    End Function

    Public Function RecuperaProductosYSubProd(ByVal pdFecSist As Date) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred
        C = New ComTipos.ClsWRegistroCred

        C.cboProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        C.cboSubProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1600")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSist, "dd/MM/yyyy") & "'")
        RecuperaProductosYSubProd = C

        C = Nothing
    End Function

    '********************************************************
    'NSILVA - Activacion de Campañas por Modalidad 04/10/2011
    '********************************************************
    Public Function RecuperaProductosYSubProdCampaniasCalendario(ByVal pdFecSist As Date, ByVal nCodModalidad As Integer) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred
        C = New ComTipos.ClsWRegistroCred

        C.cboProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        C.cboSubProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1600")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSist, "dd/MM/yyyy") & "'," & nCodModalidad)
        RecuperaProductosYSubProdCampaniasCalendario = C

        C = Nothing
    End Function


    Public Function BuscaGrupos(ByVal psNomGrp As String, ByVal pnCodAge As Integer, Optional ByVal pnCodsubProd As Integer = 1) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaGrupos '" & psNomGrp & "'," & pnCodAge & "," & pnCodsubProd)
        BuscaGrupos = D
        D = Nothing
    End Function

    Public Function BuscaActividadEcon(ByVal psNomAct As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaActEcon '" & psNomAct & "'")
        BuscaActividadEcon = D
        D = Nothing
    End Function

    Public Function DevuelveCategoriaAsesor() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DBAdminChiclayo..CredW_RecuperaCatalogoCodigos 166 ")
        DevuelveCategoriaAsesor = D
        D = Nothing
    End Function

    Public Function DevuelveTipoAsesor() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 6002 ")
        DevuelveTipoAsesor = D
        D = Nothing
    End Function

    Public Sub RegistraRolAnalista(ByVal pnCodRol As Integer, ByVal pnCodAge As Integer, ByVal oData As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistroRoles", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodRol", SqlDbType.Int).Value = pnCodRol
        oComm.Parameters.Add("@oData", SqlDbType.Xml).Value = oData.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function CargaDatosPropuestaAnalista(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pdFecSis As Date) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred

        C = New ComTipos.ClsWRegistroCred
        C.Datos.CargaTabla(, , "CredW_RecuperaPropuesta  " & Trim(Str(pnCodCred)) & "," & Trim(Str(pnCodAge)))
        C.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.cboCondicion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2070")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        C.DatosRelaPers.CargaTabla(, , "CredW_RecuperaCredPersonas  " & Trim(Str(pnCodAge)) + "," & Trim(Str(pnCodCred)))
        C.CalDesParcial.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))
        C.cboPromotores.CargaTabla(, , "Cred_RecuperaPromotores " & Trim(Str(pnCodAge)))
        C.cboRecaudadores.CargaTabla(, , "Cred_RecuperaRecaudadores " & Trim(Str(pnCodAge)))
        '---vvillafranca 06-08-2012
        C.oComentarios.CargaTabla(, , "CredW_RecuperaComentariosPropuesta  " & Trim(Str(pnCodCred)) & "," & Trim(Str(pnCodAge)))
        '--
        ''---vvillafranca 05-11-2012
        'C.oMesesDoble.CargaTabla(, , "Cred_RecuperaCuotaDoble  " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))
        ''--

        CargaDatosPropuestaAnalista = C
    End Function

    Public Function CargaDatosCreditosSugeridos(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pdFecSis As Date) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred

        C = New ComTipos.ClsWRegistroCred
        C.Datos.CargaTabla(, , "CredW_RecuperaCredSugeridos  " & Trim(Str(pnCodCred)) & "," & Trim(Str(pnCodAge)))
        C.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        C.CalDesParcial.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))

        CargaDatosCreditosSugeridos = C
    End Function
    Public Function CargaDatosCreditosAprob(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pdFecSis As Date) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred

        C = New ComTipos.ClsWRegistroCred
        C.Datos.CargaTabla(, , "CredW_RecuperaCredAprob  " & Trim(Str(pnCodCred)) & "," & Trim(Str(pnCodAge)))
        C.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        C.CalDesParcial.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))

        CargaDatosCreditosAprob = C
    End Function

    Public Function CargaModalidades(ByVal pnSubProd As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaModalidad " & pnSubProd)
        CargaModalidades = D
        D = Nothing

    End Function
    Public Function DevuelveDeudaVencida(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevuelveDeudaVencida " & pnCodAge & "," & pnCodCred)
        DevuelveDeudaVencida = D
        D = Nothing
    End Function
    Public Function DevuelveDatosCallCenter(ByVal pnCodPers As Integer, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevuelveDatosCallCenter " & pnCodPers & ", " & pnCodAge & ", " & pnCodCred)
        DevuelveDatosCallCenter = D
        D = Nothing
    End Function

    Public Function DevuelveTelefonos(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevuelveFonosPersonas " & pnCodPers)
        DevuelveTelefonos = D
        D = Nothing
    End Function
    Public Function DevuelveTelefonosPersonas(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevuelveTelefonosPersona " & pnCodPers)
        DevuelveTelefonosPersonas = D
        D = Nothing

    End Function

    Public Function DevuelveDestinos() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        DevuelveDestinos = D
        D = Nothing

    End Function
    Public Function CargaDestinos(ByVal pnProd As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDestinos " & pnProd)
        CargaDestinos = D
        D = Nothing

    End Function

    Public Function CargaExcepcionesCreditos(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_ExcepcionCreditos " & pnCodPers)
        CargaExcepcionesCreditos = D
        D = Nothing

    End Function

    Public Function CargaMotivoExcepcionesCreditos() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5070")
        CargaMotivoExcepcionesCreditos = D
        D = Nothing
    End Function

    Public Sub CargaValoresProducto(ByVal nProd As Integer, ByVal nSubProd As Integer,
                                        ByVal nCodFactor As Integer, ByVal nCodVariable As Integer,
                                       ByRef oValoresProd As ComTipos.clsDataset)
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "PA_ScoreRecuperaProductoValores " & nProd & "," & nSubProd & "," & nCodFactor & "," & nCodVariable)
        oValoresProd = D

        D = Nothing

    End Sub

    Public Sub EliminaScorePonderacion(ByVal nProd As Integer, ByVal nSubProd As Integer,
                                        ByVal nCodFactor As Integer, ByVal nCodVariable As Integer,
                                       ByRef psError As String)


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("PA_ScoreEliminaPonderacion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = nProd
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = nSubProd
        oComm.Parameters.Add("@nCodFactor", SqlDbType.Int).Value = nCodFactor
        oComm.Parameters.Add("@nCodvariable", SqlDbType.Int).Value = nCodVariable
        oComm.Parameters.Add("@psError", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        psError = oComm.Parameters("@psError").Value

        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Sub EliminaScoreProductoValores(ByVal nIDValor As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("PA_ScoreEliminaValor", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nIDValor", SqlDbType.Int).Value = nIDValor

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing


    End Sub

    Public Sub AgregaScoreProductoValores(ByVal nProd As Integer, ByVal nSubProd As Integer,
                                        ByVal nCodFactor As Integer, ByVal nCodVariable As Integer,
                                        ByVal sValor As String, ByVal nRangoIni As Double, ByVal nRangoFin As Double,
                                          ByVal nPuntaje As Integer, ByVal sNomValor As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("PA_ScoreAgregaValor", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = nProd
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = nSubProd
        oComm.Parameters.Add("@nCodFactor", SqlDbType.Int).Value = nCodFactor
        oComm.Parameters.Add("@nCodvariable", SqlDbType.Int).Value = nCodVariable
        oComm.Parameters.Add("@cValor", SqlDbType.VarChar, 100).Value = sValor
        oComm.Parameters.Add("@cNomValor", SqlDbType.VarChar, 100).Value = sNomValor
        oComm.Parameters.Add("@nRangoIni", SqlDbType.Money).Value = nRangoIni
        oComm.Parameters.Add("@nRangoFin", SqlDbType.Money).Value = nRangoFin
        oComm.Parameters.Add("@nPuntaje", SqlDbType.Int).Value = nPuntaje

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing




    End Sub

    Public Sub AgregaScorePonderaciones(ByVal nProd As Integer, ByVal nSubProd As Integer,
                                        ByVal nCodFactor As Integer, ByVal nCodVariable As Integer,
                                        ByVal nPondFactor As Double, ByVal nPondVariable As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("PA_ScoreAgregaPonderacion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = nProd
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = nSubProd
        oComm.Parameters.Add("@nCodFactor", SqlDbType.Int).Value = nCodFactor
        oComm.Parameters.Add("@nCodvariable", SqlDbType.Int).Value = nCodVariable
        oComm.Parameters.Add("@nPondFactor", SqlDbType.Money).Value = nPondFactor
        oComm.Parameters.Add("@nPondVariable", SqlDbType.Money).Value = nPondVariable

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing


    End Sub
    Public Sub CargaScoreVariables(ByVal nCodFactor As Integer, ByRef oVariables As ComTipos.clsDataset)
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "PA_CargaScoreVariables " & nCodFactor)
        oVariables = D

        D = Nothing
    End Sub

    Public Sub CargaValoresLista(ByVal nCodFactor As Integer, ByVal nCodVariable As Integer, ByRef oValoesLista As ComTipos.clsDataset)
        Dim D As ComTipos.clsDataset

        'D = New ComTipos.clsDataset
        'D.CargaTabla(, , "PA_ScoreCargaListaVariable " & nCodFactor & "," & nCodVariable)
        'oValoesLista = D

        'SEXO
        If nCodFactor = 1 And nCodVariable = 2 Then
            D = New ComTipos.clsDataset
            D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2030")
            oValoesLista = D
            D = Nothing
        End If

        'ESTADO CIVIL
        If nCodFactor = 1 And nCodVariable = 3 Then
            D = New ComTipos.clsDataset
            D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 100")
            oValoesLista = D
            D = Nothing
        End If


        'NIVEL DE EDUCACIÓN
        If nCodFactor = 1 And nCodVariable = 5 Then
            D = New ComTipos.clsDataset
            D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 3003")
            oValoesLista = D
            D = Nothing
        End If

        'TIPO DE VIVIENDA
        If nCodFactor = 1 And nCodVariable = 6 Then
            D = New ComTipos.clsDataset
            D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2158")
            oValoesLista = D
            D = Nothing
        End If

        'CALIFICACIÓN EN CENTRAL DE RIESGO
        If nCodFactor = 2 And nCodVariable = 2 Then
            D = New ComTipos.clsDataset
            D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 4550")
            oValoesLista = D
            D = Nothing
        End If


        'FRECUENCIA DE INGRESOS
        If nCodFactor = 1 And nCodVariable = 9 Then
            D = New ComTipos.clsDataset
            D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 3080")
            oValoesLista = D
            D = Nothing
        End If


        'REFERENCIAS
        If nCodFactor = 1 And nCodVariable = 10 Then
            D = New ComTipos.clsDataset
            D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 3081")
            oValoesLista = D
            D = Nothing
        End If

        D = Nothing
    End Sub

    Public Sub CargaPonderacionesProducto(ByRef oPonderaciones As ComTipos.clsDataset, ByVal pnProd As Integer, ByVal pnSubProd As Integer)
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "PA_CargaPonderaciones " & pnProd & "," & pnSubProd)
        oPonderaciones = D

        D = Nothing
    End Sub
    Public Sub CargaCombosScore(ByRef oProductos As ComTipos.clsDataset, ByRef oFactores As ComTipos.clsDataset)
        Dim D As ComTipos.clsDataset
        Dim D2 As ComTipos.clsDataset

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        oProductos = D

        D2 = New ComTipos.clsDataset
        D2.CargaTabla(, , "PA_CargaFactores ")
        oFactores = D2

        D = Nothing
    End Sub

    Public Function CargaDatosRegistro(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pdFecSis As Date) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred

        C = New ComTipos.ClsWRegistroCred
        C.Datos.CargaTabla(, , "CredW_RecuperaRegistroCred  " & Trim(Str(pnCodCred)) & "," & Trim(Str(pnCodAge)))
        C.cboAnalista.CargaTabla(, , "CredW_RecuperaAnalista  " & Trim(Str(pnCodAge)))
        C.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        'C.cboDestino.CargaTabla(, , "Cred_RecuperaModalidad 2050")
        C.cboProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        C.cboSubProd.CargaTabla(, , "Cred_Recupera_SubProd 1")
        C.DatosRelaPers.CargaTabla(, , "CredW_RecuperaCredPersonas  " & Trim(Str(pnCodAge)) + "," & Trim(Str(pnCodCred)))
        C.cboCondicion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2070")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        C.cboOficinas.CargaTabla(, , "Gen_RecuperaOficinasPorAgencia  " & Trim(Str(pnCodAge)))
        C.CalDesParcial.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))
        C.cboPromotores.CargaTabla(, , "Cred_RecuperaPromotores " & Trim(Str(pnCodAge)))
        C.cboRecaudadores.CargaTabla(, , "Cred_RecuperaRecaudadores " & Trim(Str(pnCodAge)))
        C.cboProd.DevuelveXMLLinQ()

        CargaDatosRegistro = C
    End Function

    Public Function CargaCreditosAprob(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pdFecSis As Date) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred

        C = New ComTipos.ClsWRegistroCred
        C.Datos.CargaTabla(, , "CredW_RecuperaCredAprob  " & Trim(Str(pnCodCred)) & "," & Trim(Str(pnCodAge)))
        C.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        'C.cboDestino.CargaTabla(, , "Cred_RecuperaModalidad 2050")
        C.cboProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        C.cboSubProd.CargaTabla(, , "Cred_Recupera_SubProd 1")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        C.CalDesParcial.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))

        CargaCreditosAprob = C
    End Function
    Public Function CargaPersonasRelacionadas(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset

        C.CargaTabla(, , "CredW_RecuperaCredPersonas  " & Trim(Str(pnCodAge)) + "," & Trim(Str(pnCodCred)))
        CargaPersonasRelacionadas = C
    End Function

    Public Function DevuelveMicroSeguros(ByVal pcPoliza As String, ByVal pcSolicitud As String) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset

        C.CargaTabla(, , "Cred_RecuperaMicroSeguros '" & pcPoliza & "', '" & pcSolicitud & "'")
        DevuelveMicroSeguros = C
    End Function

    Public Function DevuelveTiposPlanMicroSeguros() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset

        C.CargaTabla(, , "Cred_TiposMicroSeguros ")
        DevuelveTiposPlanMicroSeguros = C
    End Function

    Public Function DevuelveTiposCredito() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset

        C.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        DevuelveTiposCredito = C
    End Function

    Public Function DevuelveTiposDocumentos() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset

        C.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1400")
        DevuelveTiposDocumentos = C
    End Function

    Public Function CargaDatosAprobacion(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal psCodUsu As String) As ComTipos.ClsWAprobCred
        Dim C As ComTipos.ClsWAprobCred

        C = New ComTipos.ClsWAprobCred
        C.Datos.CargaTabla(, , "CredW_RecuperaParaAprobacionCredito  " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)) & ",'" & psCodUsu & "'")
        C.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.cboProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        C.cboSubProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1600")
        C.DatosRelaPers.CargaTabla(, , "CredW_RecuperaCredPersonas  " & Trim(Str(pnCodAge)) + "," & Trim(Str(pnCodCred)))
        C.cboCondicion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2070")
        C.CalDesParcial.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))
        ''---vvillafranca 05-11-2012
        'C.oMesesDoble.CargaTabla(, , "Cred_RecuperaCuotaDoble  " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))
        ''--
        CargaDatosAprobacion = C
    End Function

    Public Function RecuperaSoloTiposRelaciones() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredW_RecuperaConceptosRelaPers")
        RecuperaSoloTiposRelaciones = C
        C = Nothing
    End Function

    Public Function RecuperaMeses() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 4044")
        RecuperaMeses = C
        C = Nothing
    End Function
    'RecuperaLineasCredConsulta
    Public Function RecuperaLineasCreditoConsulta(ByVal pnCodAge As Integer, ByVal pnPlazo As Integer, ByVal pnMonto As Double, _
             ByVal pnNumPrestamo As Integer, ByVal pnMoneda As Integer, ByVal pnProd As Integer, ByVal pnSubProd As Integer, ByVal pnCodCamp As Integer, ByVal pnCodAutonom As Integer) As ComTipos.clsDataset


        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_RecuperaLineasCredConsulta " & pnCodAge & "," & pnPlazo & "," _
            & pnMonto & "," & pnNumPrestamo & "," & pnMoneda & "," & pnProd & "," & pnSubProd & "," & pnCodCamp & "," & pnCodAutonom)

        RecuperaLineasCreditoConsulta = C
        C = Nothing
    End Function

    Public Function RecuperaLineasCredito(ByVal pnCodAge As Integer, ByVal pnPlazo As Integer, ByVal pnMonto As Double, _
                ByVal pnNumPrestamo As Integer, ByVal pnMoneda As Integer, ByVal pnProd As Integer, ByVal pnSubProd As Integer, ByVal pnCodCamp As Integer) As ComTipos.clsDataset


        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredW_RecuperaLineasCred " & pnCodAge & "," & pnPlazo & "," _
            & pnMonto & "," & pnNumPrestamo & "," & pnMoneda & "," & pnProd & "," & pnSubProd & "," & pnCodCamp)

        RecuperaLineasCredito = C
        C = Nothing
    End Function

    Public Function RecuperaPorcentajesVertical() As ComTipos.clsDataset
        Dim oData As ComTipos.clsDataset
        oData = New ComTipos.clsDataset

        oData.CargaTabla(, , "")

        RecuperaPorcentajesVertical = oData
        oData = Nothing
    End Function

    Public Function RecuperaDiasMinimosPagoVertical() As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RecuperaDiasMinimosPagoVertical", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nValor", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        RecuperaDiasMinimosPagoVertical = oComm.Parameters(0).Value


        oComm = Nothing
        objCon = Nothing



    End Function

    'eduenas modificado 04/10/2012 --> modificado para pase version 23/11/2012
    Public Function RecuperaMontosNotaCreditoCancelCuotas(ByVal pnCuotasAPagar As Integer, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnNivel As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_NotaCredCancelCuotasRecuperaMontos " & pnCodAge & "," & pnCodCred & "," & pnCuotasAPagar)
        'D.CargaTabla(, , "Cred_NotaCredCancelCuotasRecuperaMontos_BK " & pnCodAge & "," & pnCodCred & "," & pnCuotasAPagar & ", " & pnNivel)
        RecuperaMontosNotaCreditoCancelCuotas = D
        D = Nothing
    End Function

    'eduenas modificado 04/10/2012 --> modificado para pase version 23/11/2012
    Public Function RecuperaMontosNotaCreditoCancelTotal(ByVal pnNivel As Integer, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_NotaCredCancelTotalRecuperaMontos " & pnCodAge & "," & pnCodCred & "," & pnNivel)
        'D.CargaTabla(, , "Cred_NotaCredCancelTotalRecuperaMontos_BK " & pnCodAge & "," & pnCodCred & "," & pnNivel)
        RecuperaMontosNotaCreditoCancelTotal = D
        D = Nothing
    End Function


    Public Function RecuperaPorcentajeVertical() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaPorcentajesVertical")
        RecuperaPorcentajeVertical = D
        D = Nothing


    End Function


    Public Function RecuperaCalendarioPendienteParaAmortizar(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim oData As ComTipos.clsDataset
        oData = New ComTipos.clsDataset

        oData.CargaTabla(, , "Cred_RecuperaCalendarioParaAmortizar " & pnCodAge & "," & pnCodCred)
        RecuperaCalendarioPendienteParaAmortizar = oData
        oData = Nothing


    End Function

    Public Function CreditoENPrejudicial(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As Boolean

        Dim oData As ComTipos.clsDataset
        oData = New ComTipos.clsDataset

        oData.CargaTabla(, , "Prejud_RecuperaDatos " & nCodAge & "," & nCodCred)

        If oData.nNumFilas > 0 Then
            CreditoENPrejudicial = True
        Else
            CreditoENPrejudicial = False
        End If
        oData = Nothing

    End Function

    Public Sub BuscarDatosCreditoCallCenter(ByVal sCodExpe As String, _
                                        ByVal sCodAge As String, _
                                        ByRef oTipoCred As ComTipos.clsCredPagoCuotas, _
                                        ByVal dFecha As Date, _
                                        ByVal nEstado As Integer)
        Dim oData As New ComTipos.clsDataset
        Dim oData2 As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante
        oData.CargaTabla(, , "CredW_BuscaExpedienteCallCenter '" & sCodAge & "','" & sCodExpe & "','" & dFecha & "'," & nEstado)
        oData2.CargaTabla(, , "DevuelveCodigosPlanilla " & sCodAge & "," & sCodExpe)
        'Return oData
        oTipoCred.Datos = oData
        oTipoCred.CodPlanilla = oData2
        oTipoCred.Bancos = oCons.RecuperaConstantes(2000)
        oData = Nothing
        oData2 = Nothing
        oCons = Nothing
    End Sub

    Public Sub BuscarDatosCredito(ByVal sCodExpe As String, _
                                        ByVal sCodAge As String, _
                                        ByRef oTipoCred As ComTipos.clsCredPagoCuotas, _
                                        ByVal dFecha As Date, _
                                        ByVal nEstado As Integer, _
                                        Optional ByVal bPregago As Boolean = False, _
                                        Optional ByVal bLibreAmortizac As Boolean = False)

        Dim oData As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante

        If bPregago Then
            oData.CargaTabla(, , "CredW_BuscaExpediente_Prepago '" & sCodAge & "','" & sCodExpe & "','" & dFecha & "'," & nEstado)
        ElseIf bLibreAmortizac Then
            oData.CargaTabla(, , "CredW_BuscaExpediente_LibreAmortizac '" & sCodAge & "','" & sCodExpe & "','" & dFecha & "'," & nEstado)
        Else
            oData.CargaTabla(, , "CredW_BuscaExpediente '" & sCodAge & "','" & sCodExpe & "','" & dFecha & "'," & nEstado)
        End If
        'Return oData
        oTipoCred.Datos = oData
        oTipoCred.Bancos = oCons.RecuperaConstantes(2000)
        oData = Nothing
        oCons = Nothing
    End Sub
    Public Function BuscarMicroSeguros_x_Cliente_Varios(ByVal sCliente As String, _
                                                ByVal nEstado As Integer, ByVal pnTipoBusq As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        If pnTipoBusq = 1 Then
            oData.CargaTabla(, , "Cred_BuscaMicroSeguros_x_Cliente_Varios '" & sCliente & "'," & nEstado)
        End If
        If pnTipoBusq = 2 Then
            oData.CargaTabla(, , "Cred_MicroSeguros_x_Cliente_Varios_DNI '" & sCliente & "'," & nEstado)
        End If
        If pnTipoBusq = 3 Then
            oData.CargaTabla(, , "Cred_BuscaMicroSeguros_x_Cliente_Varios_Codigo '" & sCliente & "'," & nEstado)
        End If

        Return oData
    End Function

    Public Function BuscarCreditos_x_Cliente_Varios(ByVal sCliente As String, _
                                                ByVal nEstado As Integer, ByVal pnTipoBusq As Integer, _
                                                ByVal bExcluyeGrupos As Boolean) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        If pnTipoBusq = 1 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios '" & sCliente & "'," & nEstado & "," & IIf(bExcluyeGrupos, 1, 0))
        End If
        If pnTipoBusq = 2 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_DNI '" & sCliente & "'," & nEstado & "," & IIf(bExcluyeGrupos, 1, 0))
        End If
        If pnTipoBusq = 3 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_Codigo '" & sCliente & "'," & nEstado & "," & IIf(bExcluyeGrupos, 1, 0))
        End If

        Return oData
    End Function

    Public Function BuscarCreditos_x_Cliente(ByVal sCliente As String, _
                                            ByVal nEstado As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente '" & sCliente & "'," & nEstado)
        Return oData
    End Function

    Public Function RegistrarAprobacionCreditoGrupal(ByVal D As ComTipos.ClsWCredSolGrp, ByVal pnValorItf As Double) As Integer()
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim Reg() As Integer
        Dim i As Integer
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset

        nMontoGasto = ObtieneGastoTotalCal(CDbl(D.Integrantes.ObtenerDato("nPrestamo")), 1, CInt(D.Datos.ObtenerDato("nProd")), CInt(D.Datos.ObtenerDato("nSubProd")))
        oFeriado = RecuperaFeriados(CDate(D.Datos.ObtenerDato("dFecDes")))

        D.Integrantes.Inicio()

        ReDim Reg(D.Integrantes.nNumFilas)
        i = 1
        Do While Not D.Integrantes.EOF

            oCal = New ComTipos.ClsCalendario
            oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(D.Datos.ObtenerDato("nCodAge")), CInt(D.Integrantes.ObtenerDato("nCodCred")), _
                2, CDbl(D.Integrantes.ObtenerDato("nPrestamo")), CInt(D.Datos.ObtenerDato("nCuotas")), CDbl(D.Integrantes.ObtenerDato("nTasaComp")), _
                CInt(D.Datos.ObtenerDato("nDiaVen")), CDate(D.Datos.ObtenerDato("dFecDes")), CInt(D.Datos.ObtenerDato("nDiasGra")), CDbl(D.Datos.ObtenerDato("nTipoGra")), 0, _
            CInt(D.Datos.ObtenerDato("nProd")), CInt(D.Datos.ObtenerDato("nSubProd")), IIf(CInt(D.Datos.ObtenerDato("nCamp")) = 0, True, False), CInt(D.Datos.ObtenerDato("nTipoPeriodo")), CInt(D.Datos.ObtenerDato("nPlazo")), pnValorItf)


            objCon = New ComConecta.clsConecta
            oComm = New SqlClient.SqlCommand("CredW_AprobarCredito", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
            oComm.Parameters("@nCodCred").Value = IIf(D.Integrantes.ObtenerDato("nCodCred") = Nothing, 0, D.Integrantes.ObtenerDato("nCodCred"))
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
            oComm.Parameters("@nCodAge").Value = IIf(D.Datos.ObtenerDato("nCodAge") = Nothing, 0, D.Datos.ObtenerDato("nCodAge"))
            oComm.Parameters.Add("@nProd", SqlDbType.Int)
            oComm.Parameters("@nProd").Value = D.Datos.ObtenerDato("nProd")
            oComm.Parameters.Add("@nSubProd", SqlDbType.Int)
            oComm.Parameters("@nSubProd").Value = D.Datos.ObtenerDato("nSubProd")
            oComm.Parameters.Add("@nPrestamo", SqlDbType.Money)
            oComm.Parameters("@nPrestamo").Value = D.Integrantes.ObtenerDato("nPrestamo")
            oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
            oComm.Parameters("@nCuotas").Value = D.Datos.ObtenerDato("nCuotas")
            oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
            oComm.Parameters("@nPlazo").Value = D.Datos.ObtenerDato("nPlazo")
            oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
            oComm.Parameters("@nDiaFijo").Value = D.Datos.ObtenerDato("nDiaVen")
            oComm.Parameters.Add("@nCodLinea", SqlDbType.Int)
            oComm.Parameters("@nCodLinea").Value = D.Integrantes.ObtenerDato("nLinea")
            oComm.Parameters.Add("@nDestino", SqlDbType.Int)
            oComm.Parameters("@nDestino").Value = D.Datos.ObtenerDato("nDestino")
            oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
            oComm.Parameters("@nTasaComp").Value = D.Integrantes.ObtenerDato("nTasaComp")
            oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money)
            oComm.Parameters("@nMontoCuota").Value = D.Integrantes.ObtenerDato("nMontoCuota")
            oComm.Parameters.Add("@nCondicion", SqlDbType.Int)
            oComm.Parameters("@nCondicion").Value = D.Datos.ObtenerDato("nCond")
            oComm.Parameters.Add("@nGastos", SqlDbType.Money)
            oComm.Parameters("@nGastos").Value = 0
            oComm.Parameters.Add("@dFecDes", SqlDbType.DateTime)
            oComm.Parameters("@dFecDes").Value = D.Datos.ObtenerDato("dFecDes")
            oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
            oComm.Parameters("@nDiasGracia").Value = CInt(D.Datos.ObtenerDato("nDiasGra"))
            'ARCV 08-12-2007
            oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
            oComm.Parameters("@nTipoGracia").Value = CInt(D.Datos.ObtenerDato("nTipoGra"))
            '---
            oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.Int)
            oComm.Parameters("@nTipoPeriodo").Value = CInt(D.Datos.ObtenerDato("nTipoPeriodo"))
            oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
            oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML

            oComm.ExecuteNonQuery()

            objCon = Nothing
            oComm = Nothing
            oCal = Nothing
            oCalData = Nothing

            Reg(i) = CInt(D.Integrantes.ObtenerDato("nCodCred"))
            i = i + 1
            D.Integrantes.Siguiente()
        Loop
        RegistrarAprobacionCreditoGrupal = Reg

    End Function


    Public Function RegistrarAprobacionCredito(ByVal D As ComTipos.clsDataset, ByVal Rela As ComTipos.clsDataset, ByVal DCal As ComTipos.clsDataset, ByVal pnItfValor As Double, ByVal psCodUsu As String, ByVal psComenAprob As String, ByVal poCredVigentes As ComTipos.clsDataset, Optional ByVal pbCuotaDoble As Boolean = False, Optional ByVal bCDPrimeraGratif As Boolean = True) As Integer ' vvillafranca Mesesdobles 05-11-2012

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset

        Dim nTasaDesgravamen As Double

        nMontoGasto = ObtieneGastoTotalCal(CDbl(D.ObtenerDato("nPrestamo")), 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))
        oFeriado = RecuperaFeriados(CDate(D.ObtenerDato("dFecDes")))

        'ARCV 21-05-2009
        nTasaDesgravamen = RecuperaTasaSeguroDesgravamen(CInt(D.ObtenerDato("nCodCred")), CInt(D.ObtenerDato("nCodAge")), _
                                    CDbl(D.ObtenerDato("nPrestamo")), 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))

        oCal = New ComTipos.ClsCalendario
        oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(D.ObtenerDato("nCodAge")), CInt(D.ObtenerDato("nCodCred")), _
                        1, CDbl(D.ObtenerDato("nPrestamo")), CInt(D.ObtenerDato("nCuotas")), CDbl(D.ObtenerDato("nTasaComp")), _
                        CInt(D.ObtenerDato("nDiaFijo")), CDate(D.ObtenerDato("dFecDes")), CInt(D.ObtenerDato("nDiasGracia")), CInt(D.ObtenerDato("TipoGracia")), CDbl(D.ObtenerDato("nGastos")), _
                        CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")), IIf(CInt(D.ObtenerDato("nCodCamp")) = 0, True, False), CInt(D.ObtenerDato("nTipoPeriodo")), CInt(D.ObtenerDato("nPlazo")), pnItfValor, IIf(D.ObtenerDato("nSoloIntereses") = "1", True, False), _
                         nTasaDesgravamen, CDate(D.ObtenerDato("dFecPrimPago")), pbCuotaDoble, bCDPrimeraGratif) 'vvillafranca 05-11-2012

        If poCredVigentes Is Nothing Then
            poCredVigentes = New ComTipos.clsDataset
        End If

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_AprobarCredito", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("CredW_AprobarCredito")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
        oComm.Parameters("@nCodCred").Value = IIf(D.ObtenerDato("nCodCred") = Nothing, 0, D.ObtenerDato("nCodCred"))
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
        oComm.Parameters("@nCodAge").Value = IIf(D.ObtenerDato("nCodAge") = Nothing, 0, D.ObtenerDato("nCodAge"))
        oComm.Parameters.Add("@nProd", SqlDbType.Int)
        oComm.Parameters("@nProd").Value = D.ObtenerDato("nProd")
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int)
        oComm.Parameters("@nSubProd").Value = D.ObtenerDato("nSubProd")
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money)
        oComm.Parameters("@nPrestamo").Value = D.ObtenerDato("nPrestamo")
        oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
        oComm.Parameters("@nCuotas").Value = D.ObtenerDato("nCuotas")
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
        oComm.Parameters("@nPlazo").Value = D.ObtenerDato("nPlazo")
        oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
        oComm.Parameters("@nDiaFijo").Value = D.ObtenerDato("nDiaFijo")
        oComm.Parameters.Add("@nCodLinea", SqlDbType.Int)
        oComm.Parameters("@nCodLinea").Value = D.ObtenerDato("nCodLinea")
        oComm.Parameters.Add("@nDestino", SqlDbType.Int)
        oComm.Parameters("@nDestino").Value = D.ObtenerDato("nDestino")
        oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
        oComm.Parameters("@nTasaComp").Value = D.ObtenerDato("nTasaComp")
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money)
        oComm.Parameters("@nMontoCuota").Value = D.ObtenerDato("nMontoCuota")
        oComm.Parameters.Add("@nCondicion", SqlDbType.Int)
        oComm.Parameters("@nCondicion").Value = D.ObtenerDato("nCodModalidad")
        oComm.Parameters.Add("@nGastos", SqlDbType.Money)
        oComm.Parameters("@nGastos").Value = D.ObtenerDato("nGastos")
        oComm.Parameters.Add("@dFecDes", SqlDbType.DateTime)
        oComm.Parameters("@dFecDes").Value = D.ObtenerDato("dFecDes")
        oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
        oComm.Parameters("@nDiasGracia").Value = CInt(D.ObtenerDato("nDiasGracia"))
        'ARCV 08-12-2007
        oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
        oComm.Parameters("@nTipoGracia").Value = CInt(D.ObtenerDato("TipoGracia"))
        '---
        oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
        oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML
        oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt)
        oComm.Parameters("@nTipoPeriodo").Value = CInt(D.ObtenerDato("nTipoPeriodo"))
        oComm.Parameters.Add("@bTasaPref", SqlDbType.VarChar, 20).Value = D.ObtenerDato("bTasaPref")

        oComm.Parameters.Add("@nTipoDesemb", SqlDbType.Int).Value = D.ObtenerDato("nTipoDesemb")
        oComm.Parameters.Add("@dCalPar", SqlDbType.Xml).Value = DCal.DevuelveXML
        oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.Int).Value = CInt(D.ObtenerDato("nSoloIntereses"))

        oComm.Parameters.Add("@bDeTercero", SqlDbType.Bit).Value = CInt(D.ObtenerDato("bDeTercero"))
        oComm.Parameters.Add("@pcCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@psComent", SqlDbType.VarChar, 5000).Value = psComenAprob
        oComm.Parameters.Add("@pnResulAutoriza", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@dCredVigentes", SqlDbType.Xml).Value = poCredVigentes.DevuelveXML
        oComm.Parameters.Add("@dFecPrimPago", SqlDbType.DateTime).Value = D.ObtenerDato("dFecPrimPago")


        oComm.Parameters.Add("@nCodInstitucion", SqlDbType.Int).Value = IIf(D.ObtenerDato("nCodInstitucion") = "@#err", 0, D.ObtenerDato("nCodInstitucion"))

        oComm.Parameters.Add("@cCodModular", SqlDbType.VarChar, 50).Value = D.ObtenerDato("cCodModular")

        '-----------vvillafranca 05-11-2012-------
        oComm.Parameters.Add("@bCuotaDoble", SqlDbType.Bit).Value = pbCuotaDoble
        oComm.Parameters.Add("@bCDPrimeraGratif", SqlDbType.Bit).Value = bCDPrimeraGratif
        'IIf(poMesesDobles Is Nothing, DBNull.Value, poMesesDobles.DevuelveXML)
        '------------------


        'ARCV 11-05-2011
        oComm.ExecuteNonQuery()
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        's-----------

        RegistrarAprobacionCredito = oComm.Parameters(26).Value
        objCon = Nothing
        oComm = Nothing
        oCal = Nothing
        oCalData = Nothing


    End Function

    Public Sub GrabarPagoCuotas(ByVal sCodExpe As String, _
                                      ByVal sCodAge As String, _
                                      ByVal sCodAgeTrans As String, _
                                      ByVal dFecha As Date, _
                                      ByVal nMonto As Double, _
                                      ByVal nCodOpe As Integer, _
                                      ByVal nCodBanco As Integer, _
                                  ByRef nOCAPITAL As Double, _
                                  ByRef nOINTERES As Double, _
                                  ByRef nOGASTOS As Double, _
                                  ByRef nOMORA As Double, _
                                  ByRef nOIGV As Double, _
                                  ByRef nOSALDO As Double, _
                                  ByRef nOPROXCUOTA As Integer, _
                                  ByRef nONROTRANS As Integer, _
                                  ByVal sUser As String, _
                                  ByRef sOAGEDIR As String, _
                                  ByRef sOAGETEL As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_PagoCuotas", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.Char, 4).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub BuscarExtornosPago(ByVal dFecha As Date, _
                                ByVal nCodAge As Integer, _
                                    ByRef oPag As ComTipos.clsCredExtorno, _
                                    ByVal sUser As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarExtornos '" & dFecha & "'," & nCodAge & ",'" & sUser & "'")
        oPag.Lista = oData
        oData = Nothing
    End Sub

    Public Sub BuscarExtornosPagoMicroSeguros(ByVal dFecha As Date, _
                                ByVal nCodAge As Integer, _
                                    ByRef oPag As ComTipos.clsCredExtorno, _
                                    ByVal sUser As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarExtornosMicroSeguros '" & dFecha & "'," & nCodAge & ",'" & sUser & "'")
        oPag.Lista = oData
        oData = Nothing
    End Sub

    Public Sub ExtornarPago(ByVal nNroTrans As Integer, _
                        Optional ByVal sUser As String = "", _
                        Optional ByVal sComentario As String = "")
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_ExtornoPago", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTrans
        oComm.Parameters.Add("@cUserExt", SqlDbType.VarChar, 50).Value = sUser
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 400).Value = sComentario

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ExtornarPagoMicroSeguros(ByVal nNroTrans As Integer, _
                        Optional ByVal sUser As String = "", _
                        Optional ByVal sComentario As String = "")
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_ExtornoPagoMicroSeguros ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTrans
        oComm.Parameters.Add("@cUserExt", SqlDbType.VarChar, 50).Value = sUser
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 400).Value = sComentario

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub



    Public Sub RegistrarExcepcionCredito(ByVal pnCodPers As Integer, ByVal pcUser As String, ByVal dFecReg As Date, ByVal bestado As Boolean, ByVal pcComentario As String, ByVal pnMotivo As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RegistraExcepionCreditos", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = pcUser
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@bEstado", SqlDbType.Bit).Value = bestado
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 250).Value = pcComentario
        oComm.Parameters.Add("@nMotivo", SqlDbType.Int).Value = pnMotivo

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Function RecuperaCreditosPorCancelar(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset

        oData.CargaTabla(, , "Cred_RecuperaCreditosParaCancelar '" & pnCodAge & "','" & pnCodCred & "'")
        RecuperaCreditosPorCancelar = oData

        oData = Nothing
    End Function

    Public Sub BuscarDesembolso(ByVal sCodExpe As String, _
                                        ByVal sCodAge As String, _
                                        ByRef oDesemb As ComTipos.clsCredDesembolso)
        Dim oData As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante

        oData.CargaTabla(, , "CredW_BuscaDesembolso '" & sCodAge & "','" & sCodExpe & "'")
        oDesemb.Datos = oData
        oDesemb.Bancos = oCons.RecuperaConstantes(2000)

        oData = Nothing
    End Sub
    Public Function BuscarAgenciaBN() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset

        oData.CargaTabla(, , "Recupera_AgenciaBN")

        BuscarAgenciaBN = oData
        oData = Nothing
    End Function

    Public Sub DesembolsarCredito(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByVal sCodAgeTrans As String, _
                                    ByVal dFecha As Date, _
                                    ByVal nPrestamo As Double, _
                                    ByVal nCodOpe As Integer, _
                                    ByVal nCodBanco As Integer, _
                                    ByVal nCuotas As Integer, _
                                    ByVal nTasa As Double, _
                                    ByVal nDia As Integer, _
                                    ByVal nDiasGracia As Integer, _
                                    ByVal nTipoGracia As Integer, _
                                    ByVal nGastos As Double, _
                                ByRef dOFecVenc As Date, _
                                ByRef cOPagare As String, _
                                ByRef cODocEmi As String, _
                                ByRef cODomEmi As String, _
                                ByRef cOConyuge As String, _
                                ByRef cODocCon As String, _
                                ByRef cODomCon As String, _
                            ByRef cODpto As String, _
                            ByRef cOProv As String, _
                            ByRef cODist As String, _
                            ByRef cOLinea As String, _
                            ByRef cOMoneda As String, _
                            ByRef bODNI As Boolean, _
                            ByVal sUser As String, _
                            ByRef cOCodBCP As String, _
                            ByRef cOCodBN As String, _
                            ByVal nProd As Integer, _
                            ByVal nSubProd As Integer, _
                            ByRef cOGrupo As String, _
                            ByVal nPlazo As Integer, _
                            ByVal nTipoPeriodo As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim sGastos As String
        Dim sCalend As String
        Dim oTemp As Object
        'Dim MatGastosCol As Object
        Dim MatCalendCol As Object
        Dim dsCalend As New ComTipos.clsDataset
        Dim dsGastos As New ComTipos.clsDataset
        Dim dsCamp As New ComTipos.clsDataset
        Dim oCalend As New ComTipos.ClsCalendario
        Dim i As Integer
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset

        nMontoGasto = ObtieneGastoTotalCal(nPrestamo, 1, nProd, nSubProd)
        oFeriado = RecuperaFeriados(dFecha)

        dsCamp.CargaTabla(, , "CredW_RecuperaCreditoDatos " & sCodExpe)

        oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamo, nCuotas, nTasa, nDia, dFecha, nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo)
        dsCamp = Nothing

        oCalend = Nothing

        'ReDim MatGastosCol(3)
        'MatGastosCol(1) = "nNroCuota"
        'MatGastosCol(2) = "nMonto"
        'MatGastosCol(3) = "nCodConcepto"

        'ARCV 29-10-2007 YA NO CREARLOS ACA
        'With dsGastos
        '    .CrearVacio(MatGastosCol)
        '    For i = 1 To UBound(oTemp)
        '        .AddRegistro()
        '        .Seek(i)
        '        .AsignarDato(oTemp(i, 2), "nNroCuota")
        '        .AsignarDato(oTemp(i, 7), "nMonto")
        '    Next i
        'End With

        Call dsGastos.CargaTabla(, , "CredW_RecuperaGastosAlDesembDet " & nPrestamo & "," & CInt(sCodExpe) & "," & CInt(sCodAge))

        'If dsGastos.nNumFilas = 0 Then
        'sGastos = ""
        'Else
        sGastos = dsGastos.DevuelveXML
        'End If

        dsGastos = Nothing

        ReDim MatCalendCol(7)
        MatCalendCol(1) = "nNroCuota"
        MatCalendCol(2) = "dFecVenc"
        MatCalendCol(3) = "nCapital"
        MatCalendCol(4) = "nIntComp"
        MatCalendCol(5) = "nIntMor"
        MatCalendCol(6) = "nIGV"
        MatCalendCol(7) = "nIntGra" 'ARCV 30-10-2007

        With dsCalend
            .CrearVacio(MatCalendCol)
            For i = 1 To UBound(oTemp)
                .AddRegistro()
                .Seek(i)
                .AsignarDato(oTemp(i, 2), "nNroCuota")
                .AsignarDato(CDate(oTemp(i, 1)).ToString("yyyyMMdd"), "dFecVenc")
                .AsignarDato(oTemp(i, 4), "nCapital")
                .AsignarDato(oTemp(i, 5), "nIntComp")
                .AsignarDato(0, "nIntMor")
                .AsignarDato(oTemp(i, 6), "nIGV")
                .AsignarDato(oTemp(i, 9), "nIntGra") 'ARCV 30-10-2007
            Next i
        End With

        sCalend = dsCalend.DevuelveXML

        dsCalend = Nothing

        dOFecVenc = CDate(oTemp(UBound(oTemp), 1)) 'Ultima fecha del Calendario

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_DesembolsarCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        'oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@XCalend", SqlDbType.Xml).Value = sCalend
        oComm.Parameters.Add("@XGastos", SqlDbType.Xml).Value = sGastos

        oComm.Parameters.Add("@cOPagare", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOPagare").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODocEmi", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODocEmi").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODomEmi", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cODomEmi").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOConyuge", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOConyuge").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODocCon", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODocCon").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODomCon", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cODomCon").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cODpto", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODpto").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOProv", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOProv").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODist", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODist").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOLinea", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOLinea").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOMoneda", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOMoneda").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@bODNI", SqlDbType.Bit)
        oComm.Parameters.Item("@bODNI").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.Char, 4).Value = sUser 'ARCV 29-10-2007

        'ARCV 31-10-2007
        oComm.Parameters.Add("@cOCodBCP", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOCodBCP").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOCodBN", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOCodBN").Direction = ParameterDirection.Output
        '-------------
        'ARCV 02-11-2007
        oComm.Parameters.Add("@cOGrupo", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOGrupo").Direction = ParameterDirection.Output
        '----------

        oComm.ExecuteNonQuery()

        cOPagare = Trim(oComm.Parameters.Item("@cOPagare").Value)
        cODocEmi = Trim(oComm.Parameters.Item("@cODocEmi").Value)
        cODomEmi = Trim(oComm.Parameters.Item("@cODomEmi").Value)
        cOConyuge = Trim(oComm.Parameters.Item("@cOConyuge").Value)
        cODocCon = Trim(oComm.Parameters.Item("@cODocCon").Value)
        cODomCon = Trim(oComm.Parameters.Item("@cODomCon").Value)

        cODpto = Trim(oComm.Parameters.Item("@cODpto").Value)
        cOProv = Trim(oComm.Parameters.Item("@cOProv").Value)
        cODist = Trim(oComm.Parameters.Item("@cODist").Value)
        cOLinea = Trim(oComm.Parameters.Item("@cOLinea").Value)
        cOMoneda = Trim(oComm.Parameters.Item("@cOMoneda").Value)
        bODNI = oComm.Parameters.Item("@bODNI").Value

        'ARCV 31-10-2007
        cOCodBCP = Trim(oComm.Parameters.Item("@cOCodBCP").Value)
        cOCodBN = Trim(oComm.Parameters.Item("@cOCodBN").Value)
        '-------
        cOGrupo = Trim(oComm.Parameters.Item("@cOGrupo").Value) 'ARCV 02-11-2007
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ExtornarDesembolso(ByVal nNroTrans As Integer, _
                        ByRef nResul As Integer, _
                        Optional ByVal sUser As String = "", _
                        Optional ByVal sComentario As String = "")

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_ExtornoDesembolso", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTrans
        oComm.Parameters.Add("@nResul", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cUserExt", SqlDbType.VarChar, 50).Value = sUser
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 400).Value = sComentario

        oComm.ExecuteNonQuery()

        nResul = oComm.Parameters("@nResul").Value

        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Sub BuscarExtornosDesemb(ByVal dFecha As Date, _
                                    ByVal nCodAge As Integer, _
                                    ByRef oPag As ComTipos.clsCredExtorno, _
                                    ByVal cUser As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarExtornosDesemb '" & dFecha & "'," & nCodAge & ",'" & cUser & "'")
        oPag.Lista = oData
        oData = Nothing
    End Sub

    Public Sub GrabarPagoCuotasCastigados(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByVal sCodAgeTrans As String, _
                                    ByVal dFecha As Date, _
                                    ByVal nMonto As Double, _
                                    ByVal nCodOpe As Integer, _
                                    ByVal nCodBanco As Integer, _
                                ByRef nOCAPITAL As Double, _
                                ByRef nOINTERES As Double, _
                                ByRef nOGASTOS As Double, _
                                ByRef nOMORA As Double, _
                                ByRef nOIGV As Double, _
                                ByRef nOSALDO As Double, _
                                ByRef nOPROXCUOTA As Integer, _
                                ByRef nONROTRANS As Integer, _
                                ByVal sUser As String, _
                                  ByRef sOAGEDIR As String, _
                                  ByRef sOAGETEL As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_PagoCuotasCastigados", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.Char, 4).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarCrediPago(ByRef oCreditos As ComTipos.clsDataset, _
                             ByVal pnCodBanco As Integer, _
                             ByRef psError As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta

        With oCreditos
            .Inicio()
            While Not .EOF
                oComm = New SqlCommand("CredW_GrabaCrediPago", objCon.DameConexion)
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = .ObtenerDato("CodAge")
                oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = .ObtenerDato("CodCred")
                oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = CDate(.ObtenerDato("Fecha"))
                oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = CDbl(.ObtenerDato("Monto"))
                oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = pnCodBanco

                oComm.Parameters.Add("@cERROR", SqlDbType.VarChar, 100)
                oComm.Parameters.Item("@cERROR").Direction = ParameterDirection.Output

                oComm.ExecuteNonQuery()

                psError = oComm.Parameters.Item("@cERROR").Value

                oComm = Nothing

                If psError = "MAYORDEUDA" Then
                    objCon = Nothing
                    Exit Sub
                End If
                .Siguiente()
            End While
        End With
        objCon = Nothing

    End Sub


    Public Sub GrabarCrediPagoGrupo(ByRef oCreditos As ComTipos.clsDataset, _
                                   ByVal pnCodBanco As Integer, _
                                   ByRef psError As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta

        With oCreditos
            .Inicio()
            While Not .EOF
                oComm = New SqlCommand("CredW_GrabaCrediPagoGrupo", objCon.DameConexion)
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = .ObtenerDato("CodAge")
                oComm.Parameters.Add("@cCod_Grupo", SqlDbType.VarChar).Value = .ObtenerDato("CodGrupo")
                oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = CDate(.ObtenerDato("Fecha"))
                oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = CDbl(.ObtenerDato("Monto"))
                oComm.Parameters.Add("@nTipoGrupo", SqlDbType.Int).Value = CInt(.ObtenerDato("TipoGrupo"))
                oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = pnCodBanco

                oComm.Parameters.Add("@cERROR", SqlDbType.VarChar, 100)
                oComm.Parameters.Item("@cERROR").Direction = ParameterDirection.Output

                oComm.ExecuteNonQuery()

                psError = oComm.Parameters.Item("@cERROR").Value

                oComm = Nothing

                'If psError <> "" Then
                'objCon = Nothing
                'Exit Sub
                'End If

                .Siguiente()
            End While
        End With
        objCon = Nothing

    End Sub

    Public Function RecuperaCuotaDECredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnIGV As Double) As Double
        Dim D As ComTipos.clsDataset
        Dim nCuota As Double
        Dim MatCuotas(500, 2) As Double
        Dim nCont As Integer
        Dim i As Integer
        Dim nPuntaje As Integer
        Dim nCursor, j As Integer

        nCuota = 0
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_DuplicCalend " & pnCodAge & "," & pnCodCred)
        nCont = 0
        D.Inicio()
        Do While Not D.EOF

            If nCuota <> CDbl(Format(CDbl(D.ObtenerDato("nCapital")) + CDbl(D.ObtenerDato("nIntComp")) + CDbl(D.ObtenerDato("nIntGra")) + CDbl(Format((CDbl(D.ObtenerDato("nIntComp")) + CDbl(D.ObtenerDato("nIntGra"))) * pnIGV, "#,0.00")), "#,0.00")) Then
                nCuota = CDbl(Format(CDbl(D.ObtenerDato("nCapital")) + CDbl(D.ObtenerDato("nIntComp")) + CDbl(D.ObtenerDato("nIntGra")) + CDbl(Format((CDbl(D.ObtenerDato("nIntComp")) + CDbl(D.ObtenerDato("nIntGra"))) * pnIGV, "#,0.00")), "#,0.00"))

                nCursor = -1

                For j = 0 To nCont - 1
                    If MatCuotas(j, 0) = nCuota Then
                        nCursor = j
                        Exit For
                    End If
                Next j

                If nCursor = -1 Then
                    nCont = nCont + 1
                    'ReDim Preserve MatCuotas(nCont, 2)
                    MatCuotas(nCont - 1, 0) = nCuota
                    MatCuotas(nCont - 1, 1) = 1
                    nCursor = nCont - 1
                Else
                    MatCuotas(nCursor, 1) = MatCuotas(nCursor, 1) + 1
                End If


            Else
                MatCuotas(nCursor, 1) = MatCuotas(nCursor, 1) + 1

            End If

            D.Siguiente()
        Loop
        D = Nothing

        nPuntaje = 0
        RecuperaCuotaDECredito = 0
        For i = 0 To nCont - 1
            If MatCuotas(i, 1) > nPuntaje Then
                RecuperaCuotaDECredito = MatCuotas(i, 0)
                nPuntaje = MatCuotas(i, 1)
            End If
        Next


    End Function


    Public Function RenovarDesembolso(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, _
         ByVal pdFecSis As Date, ByVal pnValorItf As Double, ByVal pnPrestamoNew As Double, ByVal pnSoloIntereses As Integer) As ComTipos.clsDataset

        Dim nSaldoK As Double
        Dim nMontoCuota As Double
        Dim C As ComTipos.clsDataset
        Dim nTasaComp As Double
        Dim nIntComp As Double
        Dim Cal As ComTipos.ClsCalendario
        Dim nDias As Integer
        Dim dFechaTemp As Date
        Dim nNrocuota As Integer
        Dim nTipoPeriodo As Integer
        Dim nDiaFijo As Integer
        Dim nPlazo As Integer
        Dim MatFechas As Object
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oFeriado As ComTipos.clsDataset
        Dim nIntAFecha As Double
        Dim nIntAFechaDECuota As Double

        Dim nCapitalTemp As Double
        Dim pnTipoRenov As Integer
        Dim pnIgv As Double
        Dim I As Integer

        If pnValorItf > 0 Then
            pnIgv = 0
        Else
            pnIgv = 0.19
        End If


        pnTipoRenov = 2

        nCapitalTemp = 0

        C = New ComTipos.clsDataset

        nMontoCuota = RecuperaCuotaDECredito(pnCodAge, pnCodCred, pnIgv)

        C.CargaTabla(, , "Cred_RecuperaDatosBasicosDECredito " & pnCodAge & "," & pnCodCred)
        nTasaComp = CDbl(C.ObtenerDato("nTasaComp"))
        nSaldoK = CDbl(C.ObtenerDato("nSaldoK")) + pnPrestamoNew
        nDiaFijo = CDbl(C.ObtenerDato("nDiaFijo"))
        nTipoPeriodo = CDbl(C.ObtenerDato("nTipoPeriodo"))
        nPlazo = CInt(C.ObtenerDato("nPlazo"))

        C = Nothing

        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_RecuperaDatosParaRenovacion " & pnCodAge & "," & pnCodCred & ",'" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        nIntAFecha = CDbl(C.ObtenerDato("nIntAFecha"))
        C = Nothing


        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_RecupeCalendarioParaRenov " & pnCodAge & "," & pnCodCred & ",'" & Format(pdFecSis, "dd/MM/yyyy") & "'")

        C.Inicio()
        Do While Not C.EOF
            C.AsignarDato("0.00", "nCapital")
            C.AsignarDato("0.00", "nIntComp")
            C.AsignarDato("0.00", "nIntGra")
            C.AsignarDato("0.00", "nIGV")
            C.AsignarDato("0", "nNroCuota")
            C.Siguiente()
        Loop

        dFechaTemp = pdFecSis
        nIntComp = 0
        nNrocuota = 0

        Cal = New ComTipos.ClsCalendario
        'SI Renovacion manteniendo la Cuota
        If pnTipoRenov = 1 Then
            I = 0
            C.Inicio()
            Do While Not C.EOF And nSaldoK > 0
                I = I + 1
                nNrocuota = nNrocuota + 1
                nDias = DateDiff(DateInterval.Day, dFechaTemp, CDate(C.ObtenerDato("dFecVenc")))
                nIntComp = Cal.CalculaInteresComp(nTasaComp, nDias, nSaldoK)

                C.AsignarDato(Format(nIntComp, "#0.00"), "nIntComp")
                C.AsignarDato(Format(nIntComp * pnIgv, "#0.00"), "nIgv")

                C.AsignarDato(Format(nMontoCuota - (CDbl(C.ObtenerDato("nIntComp")) + CDbl(C.ObtenerDato("nIgv"))), "#0.00"), "nCapital")


                If (nSaldoK - CDbl(C.ObtenerDato("nCapital"))) <= 0 Then
                    C.AsignarDato(Format(nSaldoK, "#0.00"), "nCapital")
                End If

                C.AsignarDato(nNrocuota, "nNroCuota")

                dFechaTemp = CDate(C.ObtenerDato("dFecVenc"))
                nSaldoK = nSaldoK - CDbl(C.ObtenerDato("nCapital"))

                C.Siguiente()
            Loop

            'SI ES REDUCCION DE CUOTA
        Else
            oFeriado = New ComTipos.clsDataset
            oFeriado = RecuperaFeriados(pdFecSis)
            Cal.oFeriados = oFeriado

            '************************************************
            'SI ES FECHA FIJA
            '************************************************
            'If nTipoPeriodo = 1 Then

            '    MatFechas = Cal.CalendarioFechas(pdFecSis, nDiaFijo, C.nNumFilas, 0, 0, 0)
            'Else
            '    '************************************************
            '    'SI ES PERIODO FIJO
            '    '************************************************
            '    MatFechas = Cal.CalendarioFechasCuotaFija(pdFecSis, nPlazo, C.nNumFilas, 0, 0, 0)
            'End If
            '*********************************************************************************

            C.Inicio()
            ReDim MatFechas(C.nNumFilas)
            For I = 1 To C.nNumFilas
                MatFechas(I) = C.ObtenerDato("dFecVenc")
                C.Siguiente()
            Next

            '****************************************************************************************

            nMontoCuota = Cal.CuotaFechaFija(pdFecSis, nTasaComp, pnIgv * 100, 0, C.nNumFilas, MatFechas, nSaldoK, 0)
            nMontoCuota = CDbl(Format(nMontoCuota, "#,0.00"))

            C.Inicio()
            I = 0
            nIntAFechaDECuota = CDbl(Format((nIntAFecha / C.nNumFilas), "#,0.00"))

            Do While Not C.EOF And nSaldoK > 0
                I = I + 1
                nNrocuota = nNrocuota + 1
                nDias = DateDiff(DateInterval.Day, dFechaTemp, CDate(C.ObtenerDato("dFecVenc")))
                nIntComp = Cal.CalculaInteresComp(nTasaComp, nDias, nSaldoK)

                C.AsignarDato(Format(nIntComp, "#0.00"), "nIntComp")

                If pnValorItf = 0 Then
                    C.AsignarDato(Format((nIntComp + nIntAFechaDECuota) * pnIgv, "#0.00"), "nIgv")
                Else
                    C.AsignarDato("0.00", "nIgv")
                End If

                If pnSoloIntereses = 1 Then
                    If I = C.nNumFilas Then
                        C.AsignarDato(Format(nSaldoK, "#0.00"), "nCapital")
                    Else
                        C.AsignarDato("0.00", "nCapital")
                    End If

                Else
                    C.AsignarDato(Format(nMontoCuota - (CDbl(C.ObtenerDato("nIntComp")) + CDbl(C.ObtenerDato("nIgv"))), "#0.00"), "nCapital")
                End If

                If (nSaldoK - CDbl(C.ObtenerDato("nCapital"))) <= 0 Then
                    C.AsignarDato(Format(nSaldoK, "#0.00"), "nCapital")
                End If

                C.AsignarDato(Format(nIntComp + nIntAFechaDECuota, "#0.00"), "nIntComp")

                C.AsignarDato(nNrocuota, "nNroCuota")

                dFechaTemp = CDate(C.ObtenerDato("dFecVenc"))
                nSaldoK = nSaldoK - CDbl(C.ObtenerDato("nCapital"))

                nCapitalTemp = CDbl(C.ObtenerDato("nCapital"))

                If C.nFilaActual = C.nNumFilas And nSaldoK > 0 Then
                    C.AsignarDato(Format(nCapitalTemp + nSaldoK, "#0.00"), "nCapital")
                End If

                C.Siguiente()

            Loop

        End If

        C.Inicio()
        Do While Not C.EOF
            C.AsignarDato(Format(CDate(C.ObtenerDato("dFecVenc")), "yyyy/MM/dd"), "dFecVenc")

            C.Siguiente()
        Loop


        RenovarDesembolso = C

        ''objCon = New ComConecta.clsConecta
        ''oComm = New SqlClient.SqlCommand("Cred_RenuevaCalendario", objCon.DameConexion)
        ''oComm.CommandType = CommandType.StoredProcedure
        ''oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Direction = ParameterDirection.Input
        ''oComm.Parameters("@pnCodCred").Value = pnCodCred
        ''oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Direction = ParameterDirection.Input
        ''oComm.Parameters("@pnCodAge").Value = pnCodAge
        ''oComm.Parameters.Add("@pCalen", SqlDbType.Xml).Direction = ParameterDirection.Input
        ''oComm.Parameters("@pCalen").Value = C.DevuelveXML

        ''oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
        oFeriado = Nothing

        Cal = Nothing
        C = Nothing

    End Function

    Public Sub RenovarCalendario(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, _
        ByVal pnTipoRenov As Integer, ByVal pdFecSis As Date, ByVal pnIgv As Double, ByVal pnValorItf As Double)

        Dim nSaldoK As Double
        Dim nMontoCuota As Double
        Dim C As ComTipos.clsDataset
        Dim nTasaComp As Double
        Dim nIntComp As Double
        Dim Cal As ComTipos.ClsCalendario
        Dim nDias As Integer
        Dim dFechaTemp As Date
        Dim nNrocuota As Integer
        Dim nTipoPeriodo As Integer
        Dim nDiaFijo As Integer
        Dim nPlazo As Integer
        Dim MatFechas As Object
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oFeriado As ComTipos.clsDataset
        Dim nCapitalTemp As Double
        Dim pdFechaPrimerPago As Date
        Dim pnTasaSegDesgravamen As Double
        Dim pnProd As Integer
        Dim pnSubProd As Integer
        Dim MatCalend As Object
        Dim nTieneSegDesgravamen As Integer

        nCapitalTemp = 0

        C = New ComTipos.clsDataset

        nMontoCuota = RecuperaCuotaDECredito(pnCodAge, pnCodCred, pnIgv)

        C.CargaTabla(, , "Cred_RecuperaDatosBasicosDECredito " & pnCodAge & "," & pnCodCred)
        nTasaComp = CDbl(C.ObtenerDato("nTasaComp"))
        nSaldoK = CDbl(C.ObtenerDato("nSaldoK"))
        nDiaFijo = CDbl(C.ObtenerDato("nDiaFijo"))
        nTipoPeriodo = CDbl(C.ObtenerDato("nTipoPeriodo"))
        nPlazo = CInt(C.ObtenerDato("nPlazo"))
        pnTasaSegDesgravamen = CDbl(C.ObtenerDato("nTasaSegDesgravamen"))
        pnProd = CInt(C.ObtenerDato("nProd"))
        pnSubProd = CInt(C.ObtenerDato("nSubProd"))
        nTieneSegDesgravamen = CInt(IIf(UCase(C.ObtenerDato("bTieneSegDesgravamen")) = "FALSE", 0, 1))

        If nTieneSegDesgravamen = 0 Then
            pnTasaSegDesgravamen = 0
        End If

        C = Nothing

        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_RecupeCalendarioParaRenov " & pnCodAge & "," & pnCodCred & ",'" & Format(pdFecSis, "dd/MM/yyyy") & "'")

        C.Inicio()
        Do While Not C.EOF
            C.AsignarDato("0.00", "nCapital")
            C.AsignarDato("0.00", "nIntComp")
            C.AsignarDato("0.00", "nIntGra")
            C.AsignarDato("0.00", "nIGV")
            C.AsignarDato("0", "nNroCuota")
            C.Siguiente()
        Loop

        dFechaTemp = pdFecSis
        nIntComp = 0
        nNrocuota = 0

        Cal = New ComTipos.ClsCalendario
        ''SI Renovacion manteniendo la Cuota
        'If pnTipoRenov = 1 Then

        '    C.Inicio()
        '    Do While Not C.EOF And nSaldoK > 0
        '        nNrocuota = nNrocuota + 1
        '        nDias = DateDiff(DateInterval.Day, dFechaTemp, CDate(C.ObtenerDato("dFecVenc")))
        '        nIntComp = Cal.CalculaInteresComp(nTasaComp, nDias, nSaldoK)

        '        C.AsignarDato(Format(nIntComp, "#0.00"), "nIntComp")
        '        C.AsignarDato(Format(nIntComp * pnIgv, "#0.00"), "nIgv")

        '        C.AsignarDato(Format(nMontoCuota - (CDbl(C.ObtenerDato("nIntComp")) + CDbl(C.ObtenerDato("nIgv")) + CDbl(Format((nSaldoK * pnTasaSegDesgravamen / 100), "#,0.00"))), "#0.00"), "nCapital")
        '        If (nSaldoK - CDbl(C.ObtenerDato("nCapital"))) <= 0 Then
        '            C.AsignarDato(Format(nSaldoK, "#0.00"), "nCapital")
        '        End If

        '        C.AsignarDato(nNrocuota, "nNroCuota")

        '        dFechaTemp = CDate(C.ObtenerDato("dFecVenc"))
        '        nSaldoK = nSaldoK - CDbl(C.ObtenerDato("nCapital"))

        '        nCapitalTemp = CDbl(C.ObtenerDato("nCapital"))


        '        If C.nFilaActual = C.nNumFilas And nSaldoK > 0 Then
        '            C.AsignarDato(Format(nCapitalTemp + nSaldoK, "#0.00"), "nCapital")
        '        End If

        '        C.Siguiente()
        '    Loop


        '    'SI ES REDUCCION DE CUOTA
        'Else
        oFeriado = New ComTipos.clsDataset
        oFeriado = RecuperaFeriados(pdFecSis)
        Cal.oFeriados = oFeriado

        '************************************************
        'SI ES FECHA FIJA
        '************************************************
        If nTipoPeriodo = 1 Then
            'Si la diferencia entre el primer pago y la fecha del sistema es mayor a 20 entonces OK
            pdFechaPrimerPago = DateAdd(DateInterval.Month, 1, pdFecSis)
            'If DateDiff(DateInterval.Day, pdFecSis, pdFechaPrimerPago) >= 20 And _
            If DateDiff(DateInterval.Day, pdFecSis, CDate(Right("00" & Trim(Str(nDiaFijo)), 2) & "/" & Right("00" & Trim(Str(Month(pdFechaPrimerPago))), 2) & "/" & Trim(Str(Year(pdFechaPrimerPago))))) >= 20 Then 'ARCV 30-07-2012
                pdFechaPrimerPago = CDate(Right("00" & Trim(Str(nDiaFijo)), 2) & "/" & Right("00" & Trim(Str(Month(pdFechaPrimerPago))), 2) & "/" & Trim(Str(Year(pdFechaPrimerPago))))
            Else
                pdFechaPrimerPago = DateAdd(DateInterval.Month, 2, pdFecSis)
                pdFechaPrimerPago = CDate(Right("00" & Trim(Str(nDiaFijo)), 2) & "/" & Right("00" & Trim(Str(Month(pdFechaPrimerPago))), 2) & "/" & Trim(Str(Year(pdFechaPrimerPago))))
            End If


            'MatFechas = Cal.CalendarioFechas(pdFecSis, nDiaFijo, C.nNumFilas, 0, 0, 0, pdFechaPrimerPago)
        Else
            '************************************************
            'SI ES PERIODO FIJO
            '************************************************
            pdFechaPrimerPago = DateAdd(DateInterval.Day, nPlazo, pdFecSis)
            'MatFechas = Cal.CalendarioFechasCuotaFija(pdFecSis, nPlazo, C.nNumFilas, 0, 0, 0, pdFechaPrimerPago)
        End If


        MatCalend = Cal.GeneraCalendario(0, oFeriado, nSaldoK, C.nNumFilas, nTasaComp, nDiaFijo, pdFecSis, 0, 0, 0, pnProd, pnSubProd, , nTipoPeriodo, nPlazo, pnValorItf, , pnTasaSegDesgravamen, False, , , pdFechaPrimerPago)


        'nMontoCuota = Cal.CuotaFechaFija(pdFecSis, nTasaComp, pnIgv * 100, 0, C.nNumFilas, MatFechas, nSaldoK, 0, pnTasaSegDesgravamen)
        nNrocuota = 0
        C.Inicio()
        Do While Not C.EOF
            nNrocuota = nNrocuota + 1

            'Select CC.nCodAge, CC.nCodCred, CC.nNroCuota, CC.dFecVenc, ISNULL(CC.nCapital,0)nCapital, 
            '		ISNULL(CC.nIntComp,0) nIntComp,ISNULL(CC.nIntGra,0) nIntGra , ISNULL(CC.nIgv,0) nIgv

            C.AsignarDato(MatCalend(nNrocuota, 2), "nNroCuota")
            C.AsignarDato(MatCalend(nNrocuota, 1), "dFecVenc")
            C.AsignarDato(MatCalend(nNrocuota, 4), "nCapital")
            C.AsignarDato(MatCalend(nNrocuota, 5), "nIntComp")
            C.AsignarDato("0.00", "nIntGra")
            C.AsignarDato(MatCalend(nNrocuota, 6), "nIgv")
            C.AsignarDato(MatCalend(nNrocuota, 10), "nSegDesgrav")


            'nDias = DateDiff(DateInterval.Day, dFechaTemp, CDate(C.ObtenerDato("dFecVenc")))
            'nIntComp = Cal.CalculaInteresComp(nTasaComp, nDias, nSaldoK)

            'C.AsignarDato(Format(nIntComp, "#0.00"), "nIntComp")
            'C.AsignarDato(Format(nIntComp * pnIgv, "#0.00"), "nIgv")

            'C.AsignarDato(Format(nMontoCuota - (CDbl(C.ObtenerDato("nIntComp")) + CDbl(C.ObtenerDato("nIgv")) + CDbl(Format((nSaldoK * pnTasaSegDesgravamen / 100), "#,0.00"))), "#0.00"), "nCapital")
            'If (nSaldoK - CDbl(C.ObtenerDato("nCapital"))) <= 0 Then
            '    C.AsignarDato(Format(nSaldoK, "#0.00"), "nCapital")
            'End If

            'C.AsignarDato(nNrocuota, "nNroCuota")

            'dFechaTemp = CDate(C.ObtenerDato("dFecVenc"))
            'nSaldoK = nSaldoK - CDbl(C.ObtenerDato("nCapital"))

            'nCapitalTemp = CDbl(C.ObtenerDato("nCapital"))



            'If C.nFilaActual = C.nNumFilas And nSaldoK > 0 Then
            '    C.AsignarDato(Format(nCapitalTemp + nSaldoK, "#0.00"), "nCapital")
            'End If
            C.Siguiente()

        Loop




        C.Inicio()
        Do While Not C.EOF
            C.AsignarDato(Format(CDate(C.ObtenerDato("dFecVenc")), "yyyy/MM/dd"), "dFecVenc")

            C.Siguiente()
        Loop


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RenuevaCalendario", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Direction = ParameterDirection.Input
        oComm.Parameters("@pnCodCred").Value = pnCodCred
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Direction = ParameterDirection.Input
        oComm.Parameters("@pnCodAge").Value = pnCodAge
        oComm.Parameters.Add("@pCalen", SqlDbType.Xml).Direction = ParameterDirection.Input
        oComm.Parameters("@pCalen").Value = C.DevuelveXML

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
        oFeriado = Nothing

        Cal = Nothing
        C = Nothing

    End Sub

    Public Function RegistrarSolicitudGrupo(ByVal SolGrp As ComTipos.ClsWCredSolGrp, ByVal nTipoAct As Integer, ByVal pnValorItf As Double) As Integer()
        Dim Cred() As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim i As Integer
        Dim Rela As ComTipos.clsDataset
        Dim RelaCamp(5) As String
        Dim P As Integer
        Dim IntegTempo As ComTipos.clsDataset
        Dim Conyuges As ComTipos.clsDataset
        Dim S As ComTipos.ClsSerializa
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset
        Dim nTipoSegDesg As Integer

        SolGrp.Integrantes.Inicio()
        ReDim Cred(SolGrp.Integrantes.nNumFilas)
        i = 0

        'ARCV 09-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand
        '---------------------------

        Do While Not SolGrp.Integrantes.EOF
            i = i + 1
            SolGrp.Integrantes.AsignarDato(IIf(SolGrp.Integrantes.ObtenerDato("nCoord") = "true", "1", "0"), "nCoord")

            nMontoGasto = ObtieneGastoTotalCal(CDbl(SolGrp.Integrantes.ObtenerDato("nPrestamo")), 1, CInt(SolGrp.Datos.ObtenerDato("nProd")), CInt(SolGrp.Datos.ObtenerDato("nSubProd")))
            oFeriado = RecuperaFeriados(CDate(SolGrp.Datos.ObtenerDato("dFecDes")))


            oCal = New ComTipos.ClsCalendario
            oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(SolGrp.Datos.ObtenerDato("nCodAge")), IIf(nTipoAct = 2, SolGrp.Integrantes.ObtenerDato("nCodCred"), 0), _
                1, CDbl(SolGrp.Integrantes.ObtenerDato("nPrestamo")), CInt(SolGrp.Datos.ObtenerDato("nCuotas")), CDbl(SolGrp.Integrantes.ObtenerDato("nTasaComp")), _
                CInt(SolGrp.Datos.ObtenerDato("nDiaVen")), CDate(SolGrp.Datos.ObtenerDato("dFecDes")), CInt(SolGrp.Datos.ObtenerDato("nDiasGra")), CInt(SolGrp.Datos.ObtenerDato("nTipoGra")), 0, _
                CInt(SolGrp.Datos.ObtenerDato("nProd")), CInt(SolGrp.Datos.ObtenerDato("nSubProd")), IIf(CInt(SolGrp.Datos.ObtenerDato("nCamp")) = 0, True, False), CInt(SolGrp.Datos.ObtenerDato("nTipoPeriodo")), CInt(SolGrp.Datos.ObtenerDato("nPlazo")), pnValorItf)

            'ARCV 09-05-2011
            'objCon = New ComConecta.clsConecta
            'oComm = New SqlClient.SqlCommand("CredW_ActualizaSolcitudCredito", objCon.DameConexion)
            oComm.CommandText = "CredW_ActualizaSolcitudCredito"
            oComm.Parameters.Clear()
            '--------
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
            oComm.Parameters("@nCodCred").Value = IIf(nTipoAct = 2, SolGrp.Integrantes.ObtenerDato("nCodCred"), 0)

            oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
            oComm.Parameters("@nCodAge").Value = IIf(SolGrp.Datos.ObtenerDato("nCodAge") = Nothing, 0, SolGrp.Datos.ObtenerDato("nCodAge"))
            oComm.Parameters.Add("@nProd", SqlDbType.Int)
            oComm.Parameters("@nProd").Value = SolGrp.Datos.ObtenerDato("nProd")
            oComm.Parameters.Add("@nSubProd", SqlDbType.Int)
            oComm.Parameters("@nSubProd").Value = SolGrp.Datos.ObtenerDato("nSubProd")
            oComm.Parameters.Add("@nPrestamo", SqlDbType.Money)
            oComm.Parameters("@nPrestamo").Value = SolGrp.Integrantes.ObtenerDato("nPrestamo")
            oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
            oComm.Parameters("@nCuotas").Value = SolGrp.Datos.ObtenerDato("nCuotas")
            oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
            oComm.Parameters("@nPlazo").Value = SolGrp.Datos.ObtenerDato("nPlazo")
            oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
            oComm.Parameters("@nDiaFijo").Value = SolGrp.Datos.ObtenerDato("nDiaVen")
            oComm.Parameters.Add("@nGrupo", SqlDbType.Int)
            oComm.Parameters("@nGrupo").Value = SolGrp.Datos.ObtenerDato("nCodGrp")
            oComm.Parameters.Add("@nCiclo", SqlDbType.Int)
            oComm.Parameters("@nCiclo").Value = SolGrp.Datos.ObtenerDato("nCiclo")
            oComm.Parameters.Add("@nCodLinea", SqlDbType.Int)
            oComm.Parameters("@nCodLinea").Value = SolGrp.Integrantes.ObtenerDato("nLinea")
            oComm.Parameters.Add("@nDestino", SqlDbType.Int)
            oComm.Parameters("@nDestino").Value = SolGrp.Datos.ObtenerDato("nDestino")
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int)
            oComm.Parameters("@nCodPers").Value = SolGrp.Datos.ObtenerDato("nAnalista")
            oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
            oComm.Parameters("@nTasaComp").Value = SolGrp.Integrantes.ObtenerDato("nTasaComp")
            oComm.Parameters.Add("@nTasaMor", SqlDbType.Money)
            oComm.Parameters("@nTasaMor").Value = 0
            oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money)
            oComm.Parameters("@nMontoCuota").Value = SolGrp.Integrantes.ObtenerDato("nMontoCuota")
            oComm.Parameters.Add("@nCondicion", SqlDbType.Int)
            oComm.Parameters("@nCondicion").Value = SolGrp.Datos.ObtenerDato("nCodModalidad")

            'Crea las Relaciones
            '*************************************************************
            RelaCamp(1) = "Codigo"
            RelaCamp(2) = "Documento"
            RelaCamp(3) = "Cliente"
            RelaCamp(4) = "Relacion"
            RelaCamp(5) = "ID"

            Rela = New ComTipos.clsDataset
            Rela.CrearVacio(RelaCamp)
            P = Rela.AddRegistro()
            Rela.Seek(P)

            'Agrego el Titular
            Rela.AsignarDato(SolGrp.Integrantes.ObtenerDato("nCodPers"), "Codigo")
            Rela.AsignarDato("10", "ID")

            IntegTempo = New ComTipos.clsDataset
            '*****************************************
            'Clonar Objeto
            '*****************************************
            S = New ComTipos.ClsSerializa
            S.Serializar(SolGrp.Integrantes, "SolGrpTempo")
            IntegTempo = S.DesSerializar("SolGrpTempo")
            S = Nothing
            '******************************************
            IntegTempo.Inicio()
            Conyuges = New ComTipos.clsDataset
            'Agrego los Avales
            Do While Not IntegTempo.EOF
                If IntegTempo.ObtenerDato("nCodPers") <> SolGrp.Integrantes.ObtenerDato("nCodPers") Then
                    P = Rela.AddRegistro()
                    Rela.Seek(P)
                    Rela.AsignarDato(IntegTempo.ObtenerDato("nCodPers"), "Codigo")
                    Rela.AsignarDato("40", "ID")

                    'Agrego Conyuge de Aval como Aval
                    Conyuges.CargaTabla(, , "CredW_RecuperaConyuge " & IntegTempo.ObtenerDato("nCodPers"))
                    If Conyuges.nNumFilas > 0 Then
                        P = Rela.AddRegistro()
                        Rela.Seek(P)
                        Rela.AsignarDato(Conyuges.ObtenerDato("nCodPersRel"), "Codigo")
                        Rela.AsignarDato("40", "ID")
                    End If
                End If
                IntegTempo.Siguiente()
            Loop

            IntegTempo = Nothing
            IntegTempo = New ComTipos.clsDataset
            nTipoSegDesg = 1
            'Agrego el Conyuge del Titular
            IntegTempo.CargaTabla(, , "CredW_RecuperaConyuge " & SolGrp.Integrantes.ObtenerDato("nCodPers"))
            IntegTempo.Inicio()
            If IntegTempo.nNumFilas > 0 Then
                P = Rela.AddRegistro()
                Rela.Seek(P)
                Rela.AsignarDato(IntegTempo.ObtenerDato("nCodPersRel"), "Codigo")
                Rela.AsignarDato("30", "ID")
                nTipoSegDesg = 2
            End If

            oComm.Parameters.Add("@nRelaPers", SqlDbType.Xml)
            oComm.Parameters("@nRelaPers").Value = Rela.DevuelveXML()
            '************************************************************

            oComm.Parameters.Add("@nTipoAct", SqlDbType.Int)
            oComm.Parameters("@nTipoAct").Value = nTipoAct
            oComm.Parameters.Add("@nGastos", SqlDbType.Money)
            oComm.Parameters("@nGastos").Value = 0
            oComm.Parameters.Add("@dFecDes", SqlDbType.DateTime)
            oComm.Parameters("@dFecDes").Value = SolGrp.Datos.ObtenerDato("dFecDes")
            oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
            oComm.Parameters("@nDiasGracia").Value = CInt(SolGrp.Datos.ObtenerDato("nDiasGra"))
            'ARCV 08-12-2007
            oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
            oComm.Parameters("@nTipoGracia").Value = CInt(SolGrp.Datos.ObtenerDato("nTipoGra"))
            '--------

            oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
            oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML

            oComm.Parameters.Add("@nRespCreditoNew", SqlDbType.Int)
            oComm.Parameters("@nRespCreditoNew").Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCodCamp", SqlDbType.Int)
            oComm.Parameters("@nCodCamp").Value = CInt(SolGrp.Datos.ObtenerDato("nCamp"))
            oComm.Parameters.Add("@cCodAct", SqlDbType.VarChar)
            oComm.Parameters("@cCodAct").Value = SolGrp.Integrantes.ObtenerDato("nCodAct")
            oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt)
            oComm.Parameters("@nTipoPeriodo").Value = SolGrp.Datos.ObtenerDato("nTipoPeriodo")
            oComm.Parameters.Add("@nTipoSegDesg", SqlDbType.SmallInt).Value = nTipoSegDesg

            oComm.Parameters.Add("@nPorCapTrab", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nPorActFijo", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nPorCompDeuda", SqlDbType.Money).Value = 0


            'ARCV 09-05-2011
            oComm.ExecuteNonQuery()
            'objCon.EjecutaTransaction(objCon.DameConexion, oComm, IIf(i = 1, True, False))
            '----------

            Cred(i) = oComm.Parameters("@nRespCreditoNew").Value

            'ARCV 09-05-2011
            'objCon.CierraConexion()
            'objCon = Nothing
            'oComm = Nothing
            '-----------------------
            oCal = Nothing
            oCalData = Nothing

            SolGrp.Integrantes.AsignarDato(Trim(Str(Cred(i))), "nCodCred")

            'Actualiza Tabla CredGrupoExp
            'ARCV 09-05-2011
            'objCon = New ComConecta.clsConecta
            'oComm = New SqlClient.SqlCommand("CredW_RegistraCredGrupoExp", objCon.DameConexion)
            oComm.CommandText = "CredW_RegistraCredGrupoExp"
            oComm.Parameters.Clear()
            '----------------

            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int)
            oComm.Parameters("@nCodPers").Value = CInt(SolGrp.Integrantes.ObtenerDato("nCodPers"))
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
            oComm.Parameters("@nCodCred").Value = CInt(SolGrp.Integrantes.ObtenerDato("nCodCred"))
            oComm.Parameters.Add("@bCoord", SqlDbType.Bit)
            oComm.Parameters("@bCoord").Value = CBool(SolGrp.Integrantes.ObtenerDato("nCoord"))
            oComm.Parameters.Add("@nCodGrp", SqlDbType.Int)
            oComm.Parameters("@nCodGrp").Value = CInt(SolGrp.Datos.ObtenerDato("nCodGrp"))
            oComm.Parameters.Add("@nCiclo", SqlDbType.Int)
            oComm.Parameters("@nCiclo").Value = CInt(SolGrp.Datos.ObtenerDato("nCiclo"))
            oComm.Parameters.Add("@Age", SqlDbType.Int)
            oComm.Parameters("@Age").Value = CInt(SolGrp.Datos.ObtenerDato("nCodAge"))

            'ARCV 09-05-2011
            oComm.ExecuteNonQuery()
            'objCon.CierraConexion()
            'objCon = Nothing
            'oComm = Nothing
            'objCon.EjecutaTransaction(objCon.DameConexion, oComm, , IIf(i = SolGrp.Integrantes.nNumFilas, True, False))
            '-------------

            SolGrp.Integrantes.Siguiente()
        Loop

        'ARCV 09-05-2011
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
        '----------

        oCal = Nothing
        oCalData = Nothing

        RegistrarSolicitudGrupo = Cred

    End Function

    Public Function RegistrarSolicitudGrupoWIN(ByRef SolGrp As ComTipos.ClsWCredSolGrp, ByVal nTipoAct As Integer, ByVal pnValorItf As Double) As Integer()
        Dim Cred() As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim i As Integer
        Dim Rela As ComTipos.clsDataset
        Dim RelaCamp(5) As String
        Dim P As Integer
        Dim IntegTempo As ComTipos.clsDataset
        Dim Conyuges As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset
        Dim nTipoSegDesg As Integer

        SolGrp.Integrantes.Inicio()
        ReDim Cred(SolGrp.Integrantes.nNumFilas)
        i = 0

        'ARCV 09-05-2011
        'objCon = New ComConecta.clsConecta
        'oComm = New SqlClient.SqlCommand()
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------------------

        Do While Not SolGrp.Integrantes.EOF
            i = i + 1
            SolGrp.Integrantes.AsignarDato(IIf(SolGrp.Integrantes.ObtenerDato("nCoord") = "true", "1", "0"), "nCoord")

            nMontoGasto = ObtieneGastoTotalCal(CDbl(SolGrp.Integrantes.ObtenerDato("nPrestamo")), 1, CInt(SolGrp.Datos.ObtenerDato("nProd")), CInt(SolGrp.Datos.ObtenerDato("nSubProd")))
            oFeriado = RecuperaFeriados(CDate(SolGrp.Datos.ObtenerDato("dFecDes")))

            oCal = New ComTipos.ClsCalendario
            oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(SolGrp.Datos.ObtenerDato("nCodAge")), IIf(nTipoAct = 2, SolGrp.Integrantes.ObtenerDato("nCodCred"), 0),
                1, CDbl(SolGrp.Integrantes.ObtenerDato("nPrestamo")), CInt(SolGrp.Datos.ObtenerDato("nCuotas")), CDbl(SolGrp.Integrantes.ObtenerDato("nTasaComp")),
                CInt(SolGrp.Datos.ObtenerDato("nDiaVen")), CDate(SolGrp.Datos.ObtenerDato("dFecDes")), CInt(SolGrp.Datos.ObtenerDato("nDiasGra")), CInt(SolGrp.Datos.ObtenerDato("nTipoGra")), 0,
                CInt(SolGrp.Datos.ObtenerDato("nProd")), CInt(SolGrp.Datos.ObtenerDato("nSubProd")), IIf(CInt(SolGrp.Datos.ObtenerDato("nCamp")) = 0, True, False), CInt(SolGrp.Datos.ObtenerDato("nTipoPeriodo")), CInt(SolGrp.Datos.ObtenerDato("nPlazo")), pnValorItf)

            'ARCV 09-'5-2011
            objCon = New ComConecta.clsConecta
            oComm = New SqlClient.SqlCommand("CredW_ActualizaSolcitudCredito", objCon.DameConexion)
            'oComm.CommandText = "CredW_ActualizaSolcitudCredito"
            oComm.Parameters.Clear()
            '--------------
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
            oComm.Parameters("@nCodCred").Value = IIf(nTipoAct = 2, SolGrp.Integrantes.ObtenerDato("nCodCred"), 0)

            oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
            oComm.Parameters("@nCodAge").Value = IIf(SolGrp.Datos.ObtenerDato("nCodAge") = Nothing, 0, SolGrp.Datos.ObtenerDato("nCodAge"))
            oComm.Parameters.Add("@nProd", SqlDbType.Int)
            oComm.Parameters("@nProd").Value = SolGrp.Datos.ObtenerDato("nProd")
            oComm.Parameters.Add("@nSubProd", SqlDbType.Int)
            oComm.Parameters("@nSubProd").Value = SolGrp.Datos.ObtenerDato("nSubProd")
            oComm.Parameters.Add("@nPrestamo", SqlDbType.Money)
            oComm.Parameters("@nPrestamo").Value = SolGrp.Integrantes.ObtenerDato("nPrestamo")
            oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
            oComm.Parameters("@nCuotas").Value = SolGrp.Datos.ObtenerDato("nCuotas")
            oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
            oComm.Parameters("@nPlazo").Value = SolGrp.Datos.ObtenerDato("nPlazo")
            oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
            oComm.Parameters("@nDiaFijo").Value = SolGrp.Datos.ObtenerDato("nDiaVen")
            oComm.Parameters.Add("@nGrupo", SqlDbType.Int)
            oComm.Parameters("@nGrupo").Value = SolGrp.Datos.ObtenerDato("nCodGrp")
            oComm.Parameters.Add("@nCiclo", SqlDbType.Int)
            oComm.Parameters("@nCiclo").Value = SolGrp.Datos.ObtenerDato("nCiclo")
            oComm.Parameters.Add("@nCodLinea", SqlDbType.Int)
            oComm.Parameters("@nCodLinea").Value = SolGrp.Integrantes.ObtenerDato("nLinea")
            oComm.Parameters.Add("@nDestino", SqlDbType.Int)
            oComm.Parameters("@nDestino").Value = SolGrp.Datos.ObtenerDato("nDestino")
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int)
            oComm.Parameters("@nCodPers").Value = SolGrp.Datos.ObtenerDato("nAnalista")
            oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
            oComm.Parameters("@nTasaComp").Value = SolGrp.Integrantes.ObtenerDato("nTasaComp")
            oComm.Parameters.Add("@nTasaMor", SqlDbType.Money)
            oComm.Parameters("@nTasaMor").Value = 0
            oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money)
            oComm.Parameters("@nMontoCuota").Value = SolGrp.Integrantes.ObtenerDato("nMontoCuota")
            oComm.Parameters.Add("@nCondicion", SqlDbType.Int)
            oComm.Parameters("@nCondicion").Value = SolGrp.Datos.ObtenerDato("nCodModalidad")

            oComm.Parameters.Add("@nOficina", SqlDbType.Int).Value = CInt(SolGrp.Datos.ObtenerDato("nOficina"))
            oComm.Parameters.Add("@bTasaPref", SqlDbType.VarChar, 20).Value = "False"
            oComm.Parameters.Add("@Cal_Des_Par", SqlDbType.Xml)
            oComm.Parameters("@Cal_Des_Par").Value = ""
            oComm.Parameters.Add("@nTipoDesemb", SqlDbType.Int).Value = 1
            oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.Int).Value = 0



            'Crea las Relaciones
            '*************************************************************
            RelaCamp(1) = "Codigo"
            RelaCamp(2) = "Documento"
            RelaCamp(3) = "Cliente"
            RelaCamp(4) = "Relacion"
            RelaCamp(5) = "ID"

            Rela = New ComTipos.clsDataset
            Rela.CrearVacio(RelaCamp)
            P = Rela.AddRegistro()
            Rela.Seek(P)

            'Agrego el Titular
            Rela.AsignarDato(SolGrp.Integrantes.ObtenerDato("nCodPers"), "Codigo")
            Rela.AsignarDato("10", "ID")

            IntegTempo = New ComTipos.clsDataset
            '*****************************************
            'Clonar Objeto
            '*****************************************
            'S = New ComTipos.ClsSerializa
            'S.Serializar(SolGrp.Integrantes, "SolGrpTempo")
            IntegTempo = SolGrp.Integrantes.DevolverCopia
            SolGrp.Integrantes.Seek(i)
            'S = Nothing
            '******************************************
            IntegTempo.Inicio()
            Conyuges = New ComTipos.clsDataset
            'Agrego los Avales
            Do While Not IntegTempo.EOF
                If IntegTempo.ObtenerDato("nCodPers") <> SolGrp.Integrantes.ObtenerDato("nCodPers") Then
                    P = Rela.AddRegistro()
                    Rela.Seek(P)
                    Rela.AsignarDato(IntegTempo.ObtenerDato("nCodPers"), "Codigo")
                    Rela.AsignarDato("40", "ID")

                    'Agrego Conyuge de Aval como Aval
                    Conyuges.CargaTabla(, , "CredW_RecuperaConyuge " & IntegTempo.ObtenerDato("nCodPers"))
                    If Conyuges.nNumFilas > 0 Then
                        P = Rela.AddRegistro()
                        Rela.Seek(P)
                        Rela.AsignarDato(Conyuges.ObtenerDato("nCodPersRel"), "Codigo")
                        Rela.AsignarDato("40", "ID")
                    End If
                End If
                IntegTempo.Siguiente()
            Loop

            nTipoSegDesg = 1
            IntegTempo = Nothing
            IntegTempo = New ComTipos.clsDataset
            'Agrego el Conyuge del Titular
            IntegTempo.CargaTabla(, , "CredW_RecuperaConyuge " & SolGrp.Integrantes.ObtenerDato("nCodPers"))
            IntegTempo.Inicio()
            If IntegTempo.nNumFilas > 0 Then
                P = Rela.AddRegistro()
                Rela.Seek(P)
                Rela.AsignarDato(IntegTempo.ObtenerDato("nCodPersRel"), "Codigo")
                Rela.AsignarDato("30", "ID")
                nTipoSegDesg = 2
            End If

            oComm.Parameters.Add("@nRelaPers", SqlDbType.Xml)
            oComm.Parameters("@nRelaPers").Value = Rela.DevuelveXML()
            '************************************************************

            oComm.Parameters.Add("@nTipoAct", SqlDbType.Int)
            oComm.Parameters("@nTipoAct").Value = nTipoAct
            oComm.Parameters.Add("@nGastos", SqlDbType.Money)
            oComm.Parameters("@nGastos").Value = 0
            oComm.Parameters.Add("@dFecDes", SqlDbType.DateTime)
            oComm.Parameters("@dFecDes").Value = SolGrp.Datos.ObtenerDato("dFecDes")
            oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
            oComm.Parameters("@nDiasGracia").Value = CInt(SolGrp.Datos.ObtenerDato("nDiasGra"))
            'ARCV 08-12-2007
            oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
            oComm.Parameters("@nTipoGracia").Value = CInt(SolGrp.Datos.ObtenerDato("nTipoGra"))
            '--------

            oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
            oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML

            oComm.Parameters.Add("@nRespCreditoNew", SqlDbType.Int)
            oComm.Parameters("@nRespCreditoNew").Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCodCamp", SqlDbType.Int)
            oComm.Parameters("@nCodCamp").Value = CInt(SolGrp.Datos.ObtenerDato("nCamp"))
            oComm.Parameters.Add("@cCodAct", SqlDbType.VarChar)
            oComm.Parameters("@cCodAct").Value = SolGrp.Integrantes.ObtenerDato("nCodAct")
            oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt)
            oComm.Parameters("@nTipoPeriodo").Value = SolGrp.Datos.ObtenerDato("nTipoPeriodo")
            oComm.Parameters.Add("@nMoneda", SqlDbType.SmallInt)
            oComm.Parameters("@nMoneda").Value = SolGrp.Datos.ObtenerDato("nMoneda")
            oComm.Parameters.Add("@nTipoSegDesg", SqlDbType.Int).Value = nTipoSegDesg

            oComm.Parameters.Add("@nPorCapTrab", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nPorActFijo", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nPorCompDeuda", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@dFecPrimPago", SqlDbType.DateTime).Value = CDate(oCalData.ObtenerDato("dFecVenc", 1))

            'ARCV 11-05-2011
            oComm.ExecuteNonQuery()
            'If objCon.EjecutaTransaccion(oComm) = False Then
            '    RegistrarSolicitudGrupoWIN = Cred
            '    Exit Function
            'End If
            'objCon.EjecutaTransaccion(oComm)
            '--------------------

            Cred(i) = oComm.Parameters("@nRespCreditoNew").Value

            'ARCV 09-05-2011
            objCon.CierraConexion()
            objCon = Nothing
            oComm = Nothing
            '-----------------------
            oCal = Nothing
            oCalData = Nothing

            SolGrp.Integrantes.AsignarDato(Trim(Str(Cred(i))), "nCodCred", i)

            'Actualiza Tabla CredGrupoExp
            'ARCV 09-05-2011
            objCon = New ComConecta.clsConecta
            oComm = New SqlClient.SqlCommand("CredW_RegistraCredGrupoExp", objCon.DameConexion)
            'oComm.CommandText = "CredW_RegistraCredGrupoExp"
            'oComm.Parameters.Clear()
            '----------------
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int)
            oComm.Parameters("@nCodPers").Value = CInt(SolGrp.Integrantes.ObtenerDato("nCodPers"))
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
            oComm.Parameters("@nCodCred").Value = Cred(i)
            oComm.Parameters.Add("@bCoord", SqlDbType.Bit)
            oComm.Parameters("@bCoord").Value = CBool(SolGrp.Integrantes.ObtenerDato("nCoord"))
            oComm.Parameters.Add("@nCodGrp", SqlDbType.Int)
            oComm.Parameters("@nCodGrp").Value = CInt(SolGrp.Datos.ObtenerDato("nCodGrp"))
            oComm.Parameters.Add("@nCiclo", SqlDbType.Int)
            oComm.Parameters("@nCiclo").Value = CInt(SolGrp.Datos.ObtenerDato("nCiclo"))
            oComm.Parameters.Add("@Age", SqlDbType.Int)
            oComm.Parameters("@Age").Value = CInt(SolGrp.Datos.ObtenerDato("nCodAge"))

            'ARCV 11-05-2011
            oComm.ExecuteNonQuery()
            'If objCon.EjecutaTransaccion(oComm) = False Then
            '    RegistrarSolicitudGrupoWIN = Cred
            '    Exit Function
            'End If
            'objCon.EjecutaTransaccion(oComm)
            '--------------------

            SolGrp.Integrantes.Siguiente()
        Loop

        'ARCV 09-05-2011
        'objCon.FinalizaTransaccion()
        'objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
        '----------

        oCal = Nothing
        oCalData = Nothing

        RegistrarSolicitudGrupoWIN = Cred

    End Function

    Public Function RecuperaCalendDesemParcial(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaCalenDesembParcial " & pnCodCred & "," & pnCodAge)
        RecuperaCalendDesemParcial = D
        D = Nothing

    End Function

    Public Function RegistrarAprobacionGrupoWIN(ByVal SolGrp As ComTipos.ClsWCredSolGrp, ByVal nTipoAct As Integer, ByVal pnValorItf As Double) As Integer()
        Dim Cred() As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim i As Integer
        Dim Rela As ComTipos.clsDataset
        Dim RelaCamp(5) As String
        Dim P As Integer
        Dim IntegTempo As ComTipos.clsDataset
        Dim Conyuges As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset

        Dim nTasaDesgravamen As Double
        Dim pdFecPrimerPago As Date

        SolGrp.Integrantes.Inicio()
        ReDim Cred(SolGrp.Integrantes.nNumFilas)
        i = 0

        ''ARCV 11-05-2011
        'objCon = New ComConecta.clsConecta
        'oComm = New SqlClient.SqlCommand()
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        ''---------------------------

        Do While Not SolGrp.Integrantes.EOF
            i = i + 1

            If SolGrp.Integrantes.ObtenerDato("nCoord") = "1" Then
                'SolGrp.Integrantes.AsignarDato(IIf(SolGrp.Integrantes.ObtenerDato("nCoord") = "true", "1", "0"), "nCoord")

                nMontoGasto = ObtieneGastoTotalCal(CDbl(SolGrp.Integrantes.ObtenerDato("nPrestamo")), 1, CInt(SolGrp.Datos.ObtenerDato("nProd")), CInt(SolGrp.Datos.ObtenerDato("nSubProd")))
                oFeriado = RecuperaFeriados(CDate(SolGrp.Datos.ObtenerDato("dFecDes")))

                'ARCV 21-05-2009
                nTasaDesgravamen = RecuperaTasaSeguroDesgravamen(CInt(SolGrp.Datos.ObtenerDato("nCodCred")), CInt(SolGrp.Datos.ObtenerDato("nCodAge")), _
                                            CDbl(SolGrp.Datos.ObtenerDato("nPrestamo")), 1, CInt(SolGrp.Datos.ObtenerDato("nProd")), CInt(SolGrp.Datos.ObtenerDato("nSubProd")))

                oCal = New ComTipos.ClsCalendario
                oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(SolGrp.Datos.ObtenerDato("nCodAge")), IIf(nTipoAct = 2, SolGrp.Integrantes.ObtenerDato("nCodCred"), 0), _
                    1, CDbl(SolGrp.Integrantes.ObtenerDato("nPrestamo")), CInt(SolGrp.Datos.ObtenerDato("nCuotas")), CDbl(SolGrp.Integrantes.ObtenerDato("nTasaComp")), _
                    CInt(SolGrp.Datos.ObtenerDato("nDiaFijo")), CDate(SolGrp.Datos.ObtenerDato("dFecDes")), CInt(SolGrp.Datos.ObtenerDato("nDiasGracia")), CInt(SolGrp.Datos.ObtenerDato("TipoGracia")), 0, _
                    CInt(SolGrp.Datos.ObtenerDato("nProd")), CInt(SolGrp.Datos.ObtenerDato("nSubProd")), IIf(CInt(SolGrp.Datos.ObtenerDato("nCodCamp")) = 0, True, False), CInt(SolGrp.Datos.ObtenerDato("nTipoPeriodo")), CInt(SolGrp.Datos.ObtenerDato("nPlazo")), pnValorItf, , _
                    nTasaDesgravamen, pdFecPrimerPago)

                'ARCV 11-05-2011
                objCon = New ComConecta.clsConecta
                oComm = New SqlClient.SqlCommand()
                oComm.Connection = objCon.DameConexion
                objCon.IniciaTransaccion()
                oComm.Transaction = objCon.TransaccionActual
                oComm.CommandText = "CredW_AprobarCredito"
                oComm.Parameters.Clear()
                '--------------

                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
                oComm.Parameters("@nCodCred").Value = IIf(SolGrp.Integrantes.ObtenerDato("nCodCred") = Nothing, 0, SolGrp.Integrantes.ObtenerDato("nCodCred"))
                oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
                oComm.Parameters("@nCodAge").Value = IIf(SolGrp.Datos.ObtenerDato("nCodAge") = Nothing, 0, SolGrp.Datos.ObtenerDato("nCodAge"))
                oComm.Parameters.Add("@nProd", SqlDbType.Int)
                oComm.Parameters("@nProd").Value = SolGrp.Datos.ObtenerDato("nProd")
                oComm.Parameters.Add("@nSubProd", SqlDbType.Int)
                oComm.Parameters("@nSubProd").Value = SolGrp.Datos.ObtenerDato("nSubProd")
                oComm.Parameters.Add("@nPrestamo", SqlDbType.Money)
                oComm.Parameters("@nPrestamo").Value = SolGrp.Integrantes.ObtenerDato("nPrestamo")
                oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
                oComm.Parameters("@nCuotas").Value = SolGrp.Datos.ObtenerDato("nCuotas")
                oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
                oComm.Parameters("@nPlazo").Value = SolGrp.Datos.ObtenerDato("nPlazo")
                oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
                oComm.Parameters("@nDiaFijo").Value = SolGrp.Datos.ObtenerDato("nDiaFijo")
                oComm.Parameters.Add("@nCodLinea", SqlDbType.Int)
                oComm.Parameters("@nCodLinea").Value = SolGrp.Integrantes.ObtenerDato("nLinea")
                oComm.Parameters.Add("@nDestino", SqlDbType.Int)
                oComm.Parameters("@nDestino").Value = SolGrp.Datos.ObtenerDato("nDestino")
                oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
                oComm.Parameters("@nTasaComp").Value = SolGrp.Integrantes.ObtenerDato("nTasaComp")
                oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money)
                oComm.Parameters("@nMontoCuota").Value = SolGrp.Integrantes.ObtenerDato("nMontoCuota")
                oComm.Parameters.Add("@nCondicion", SqlDbType.Int)
                oComm.Parameters("@nCondicion").Value = SolGrp.Datos.ObtenerDato("nCondCred")
                oComm.Parameters.Add("@nGastos", SqlDbType.Money)
                oComm.Parameters("@nGastos").Value = SolGrp.Datos.ObtenerDato("nGastos")
                oComm.Parameters.Add("@dFecDes", SqlDbType.DateTime)
                oComm.Parameters("@dFecDes").Value = SolGrp.Datos.ObtenerDato("dFecDes")
                oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
                oComm.Parameters("@nDiasGracia").Value = CInt(SolGrp.Datos.ObtenerDato("nDiasGracia"))
                'ARCV 08-12-2007
                oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
                oComm.Parameters("@nTipoGracia").Value = CInt(SolGrp.Datos.ObtenerDato("TipoGracia"))
                '---
                oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
                oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML
                oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt)
                oComm.Parameters("@nTipoPeriodo").Value = CInt(SolGrp.Datos.ObtenerDato("nTipoPeriodo"))
                oComm.Parameters.Add("@bTasaPref", SqlDbType.VarChar, 20).Value = "False"

                oComm.Parameters.Add("@nTipoDesemb", SqlDbType.Int).Value = 1
                oComm.Parameters.Add("@dCalPar", SqlDbType.Xml).Value = ""
                oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.Int).Value = 0
                oComm.Parameters.Add("@dFecPrimPago", SqlDbType.DateTime).Value = pdFecPrimerPago

                'ARCV 11-05-2011
                'oComm.ExecuteNonQuery()
                objCon.EjecutaTransaccion(oComm)
                '--------------------

                Cred(i) = CInt(SolGrp.Integrantes.ObtenerDato("nCodCred"))

                'ARCV 11-05-2011
                objCon.FinalizaTransaccion()
                objCon = Nothing
                oComm = Nothing
                '----------

                oCal = Nothing
                oCalData = Nothing
            End If
            SolGrp.Integrantes.Siguiente()
        Loop

        ''ARCV 11-05-2011
        'objCon.FinalizaTransaccion()
        'objCon = Nothing
        'oComm = Nothing
        ''----------

        oCal = Nothing
        oCalData = Nothing

        RegistrarAprobacionGrupoWIN = Cred

    End Function


    Public Sub RegistrarTasaAprobacion(ByVal D As ComTipos.clsDataset, ByVal pcCodUsu As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("RegistraTasaAprobacion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
        oComm.Parameters("@nCodCred").Value = IIf(D.ObtenerDato("nCodCred") = Nothing, 0, D.ObtenerDato("nCodCred"))
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
        oComm.Parameters("@nCodAge").Value = IIf(D.ObtenerDato("nCodAge") = Nothing, 0, D.ObtenerDato("nCodAge"))
        oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
        oComm.Parameters("@nTasaComp").Value = D.ObtenerDato("nTasaComp")
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = pcCodUsu
        oComm.Parameters.Add("@nCodLinea", SqlDbType.Money)
        oComm.Parameters("@nCodLinea").Value = D.ObtenerDato("nCodLinea")


        oComm.ExecuteNonQuery()

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RegistrarModificacionFecPrimerPago(ByVal D As ComTipos.clsDataset, ByVal dCalPar As ComTipos.clsDataset, ByVal pnValorItf As Double)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset

        nMontoGasto = ObtieneGastoTotalCal(CDbl(D.ObtenerDato("nPrestamo")), 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))
        oFeriado = RecuperaFeriados(CDate(D.ObtenerDato("dFecDes")))

        oCal = New ComTipos.ClsCalendario
        oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(D.ObtenerDato("nCodAge")), CInt(D.ObtenerDato("nCodCred")), _
            2, CDbl(D.ObtenerDato("nPrestamo")), CInt(D.ObtenerDato("nCuotas")), CDbl(D.ObtenerDato("nTasaComp")), _
            CInt(D.ObtenerDato("nDiaFijo")), CDate(D.ObtenerDato("dFecDes")), CInt(D.ObtenerDato("nDiasGracia")), CInt(D.ObtenerDato("TipoGracia")), CDbl(D.ObtenerDato("nGastos")), _
            CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")), IIf(CInt(D.ObtenerDato("nCodCamp")) = 0, True, False), CInt(D.ObtenerDato("nTipoPeriodo")), CInt(D.ObtenerDato("nPlazo")), pnValorItf, IIf(D.ObtenerDato("nSoloIntereses") = "@#err", False, IIf(D.ObtenerDato("nSoloIntereses") = "1", True, False)), , CDate(D.ObtenerDato("dFecPrimPago")))



        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraModificacionFechaPrimerPago", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
        oComm.Parameters("@nCodCred").Value = IIf(D.ObtenerDato("nCodCred") = Nothing, 0, D.ObtenerDato("nCodCred"))
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
        oComm.Parameters("@nCodAge").Value = IIf(D.ObtenerDato("nCodAge") = Nothing, 0, D.ObtenerDato("nCodAge"))
        oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
        oComm.Parameters("@nCuotas").Value = D.ObtenerDato("nCuotas")
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
        oComm.Parameters("@nPlazo").Value = D.ObtenerDato("nPlazo")
        oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
        oComm.Parameters("@nDiaFijo").Value = D.ObtenerDato("nDiaFijo")
        oComm.Parameters.Add("@nGastos", SqlDbType.Money)
        oComm.Parameters("@nGastos").Value = D.ObtenerDato("nGastos")
        oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
        oComm.Parameters("@nDiasGracia").Value = CInt(D.ObtenerDato("nDiasGracia"))
        oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
        oComm.Parameters("@nTipoGracia").Value = CInt(D.ObtenerDato("TipoGracia"))
        oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
        oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML
        oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt)
        oComm.Parameters("@nTipoPeriodo").Value = D.ObtenerDato("nTipoPeriodo")

        oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.Int).Value = IIf(D.ObtenerDato("nSoloIntereses") = "@#err", 0, D.ObtenerDato("nSoloIntereses"))
        oComm.Parameters.Add("@nTipoSegDesg", SqlDbType.Int).Value = IIf(D.ObtenerDato("nTipoSegDesg") = "@#err", 0, D.ObtenerDato("nTipoSegDesg"))
        oComm.Parameters.Add("@pdFecPrimPago", SqlDbType.DateTime).Value = D.ObtenerDato("dFecPrimPago")

        oComm.Parameters.Add("@nCodAnalista", SqlDbType.Int).Value = D.ObtenerDato("nCodPers") 'ARCV 10-10-2011

        oComm.ExecuteNonQuery()

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
        oCal = Nothing
        oCalData = Nothing

    End Sub
    Public Sub RegistraPropuestaAnalista(ByVal D As ComTipos.clsDataset, ByVal dCalPar As ComTipos.clsDataset, ByVal pnValorItf As Double, _
ByVal dCredVig As ComTipos.clsDataset, ByVal oComentarios As String, Optional ByVal pbCuotaDoble As Boolean = False, Optional ByVal bCDPrimeraGratif As Boolean = True) ' vvillafranca Mesesdobles 05-11-2012
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset

        nMontoGasto = ObtieneGastoTotalCal(CDbl(D.ObtenerDato("nPrestamo")), 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))
        oFeriado = RecuperaFeriados(CDate(D.ObtenerDato("dFecDes")))

        oCal = New ComTipos.ClsCalendario
        oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(D.ObtenerDato("nCodAge")), CInt(D.ObtenerDato("nCodCred")), _
            2, CDbl(D.ObtenerDato("nPrestamo")), CInt(D.ObtenerDato("nCuotas")), CDbl(D.ObtenerDato("nTasaComp")), _
            CInt(D.ObtenerDato("nDiaFijo")), CDate(D.ObtenerDato("dFecDes")), CInt(D.ObtenerDato("nDiasGracia")), CInt(D.ObtenerDato("TipoGracia")), CDbl(D.ObtenerDato("nGastos")), _
            CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")), IIf(CInt(D.ObtenerDato("nCodCamp")) = 0, True, False), CInt(D.ObtenerDato("nTipoPeriodo")), CInt(D.ObtenerDato("nPlazo")), pnValorItf, IIf(D.ObtenerDato("nSoloIntereses") = "@#err", False, IIf(D.ObtenerDato("nSoloIntereses") = "1", True, False)), , CDate(D.ObtenerDato("dFecPrimPago")), pbCuotaDoble, bCDPrimeraGratif)  ' vvillafranca Mesesdobles 05-11-2012

        If dCredVig Is Nothing Then
            dCredVig = New ComTipos.clsDataset
        End If

       

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RegistraPropuestaAnalista", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("Cred_RegistraPropuestaAnalista")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
        oComm.Parameters("@nCodCred").Value = IIf(D.ObtenerDato("nCodCred") = Nothing, 0, D.ObtenerDato("nCodCred"))
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
        oComm.Parameters("@nCodAge").Value = IIf(D.ObtenerDato("nCodAge") = Nothing, 0, D.ObtenerDato("nCodAge"))
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money)
        oComm.Parameters("@nPrestamo").Value = D.ObtenerDato("nPrestamo")
        oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
        oComm.Parameters("@nCuotas").Value = D.ObtenerDato("nCuotas")
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
        oComm.Parameters("@nPlazo").Value = D.ObtenerDato("nPlazo")
        oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
        oComm.Parameters("@nDiaFijo").Value = D.ObtenerDato("nDiaFijo")
        oComm.Parameters.Add("@nCodLinea", SqlDbType.Int)
        oComm.Parameters("@nCodLinea").Value = D.ObtenerDato("nCodLinea")
        oComm.Parameters.Add("@nDestino", SqlDbType.Int)
        oComm.Parameters("@nDestino").Value = D.ObtenerDato("nDestino")
        oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
        oComm.Parameters("@nTasaComp").Value = D.ObtenerDato("nTasaComp")
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money)
        oComm.Parameters("@nMontoCuota").Value = D.ObtenerDato("nMontoCuota")
        oComm.Parameters.Add("@nCondicion", SqlDbType.Int)
        oComm.Parameters("@nCondicion").Value = D.ObtenerDato("nCodModalidad")
        oComm.Parameters.Add("@nGastos", SqlDbType.Money)
        oComm.Parameters("@nGastos").Value = D.ObtenerDato("nGastos")
        oComm.Parameters.Add("@dFecDes", SqlDbType.VarChar)
        oComm.Parameters("@dFecDes").Value = D.ObtenerDato("dFecDes")
        oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
        oComm.Parameters("@nDiasGracia").Value = CInt(D.ObtenerDato("nDiasGracia"))
        oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
        oComm.Parameters("@nTipoGracia").Value = CInt(D.ObtenerDato("TipoGracia"))
        oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
        oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML
        oComm.Parameters.Add("@nCodCamp", SqlDbType.Int)
        oComm.Parameters("@nCodCamp").Value = CInt(D.ObtenerDato("nCodCamp"))
        oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt)
        oComm.Parameters("@nTipoPeriodo").Value = D.ObtenerDato("nTipoPeriodo")
        oComm.Parameters.Add("@nMoneda", SqlDbType.SmallInt)
        oComm.Parameters("@nMoneda").Value = D.ObtenerDato("nMoneda")
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int)
        oComm.Parameters("@pnCodEval").Value = D.ObtenerDato("nCodEval")
        oComm.Parameters.Add("@cComenRatFin", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenRatFin")
        oComm.Parameters.Add("@cComenForDeb", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenForDeb")
        oComm.Parameters.Add("@cComenPlanInv", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenPlanInv")
        oComm.Parameters.Add("@cComenRieCred", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenRieCred")
        oComm.Parameters.Add("@bTasaPref", SqlDbType.VarChar, 20).Value = D.ObtenerDato("bTasaPref")

        oComm.Parameters.Add("@nTipoDesemb", SqlDbType.Int).Value = IIf(D.ObtenerDato("nTipoDesemb") = "@#err", 0, D.ObtenerDato("nTipoDesemb"))
        oComm.Parameters.Add("@dCalPar", SqlDbType.Xml).Value = IIf(dCalPar.DevuelveXML = "@#err", "", dCalPar.DevuelveXML)
        oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.Int).Value = IIf(D.ObtenerDato("nSoloIntereses") = "@#err", 0, D.ObtenerDato("nSoloIntereses"))
        oComm.Parameters.Add("@nCodPromotor", SqlDbType.Int).Value = IIf(D.ObtenerDato("nCodPromotor") = "@#err", 0, D.ObtenerDato("nCodPromotor"))
        oComm.Parameters.Add("@nCodRecaudador", SqlDbType.Int).Value = IIf(D.ObtenerDato("nCodRecaudador") = "@#err", 0, D.ObtenerDato("nCodRecaudador"))
        oComm.Parameters.Add("@nTipoSegDesg", SqlDbType.Int).Value = IIf(D.ObtenerDato("nTipoSegDesg") = "@#err", 0, D.ObtenerDato("nTipoSegDesg"))
        oComm.Parameters.Add("@pdFecPrimPago", SqlDbType.DateTime).Value = D.ObtenerDato("dFecPrimPago")
        oComm.Parameters.Add("@nCodInstitucion", SqlDbType.Int).Value = IIf(D.ObtenerDato("nCodInstitucion") = "@#err", 0, D.ObtenerDato("nCodInstitucion"))
        oComm.Parameters.Add("@cCodModular", SqlDbType.VarChar, 50).Value = D.ObtenerDato("cCodModular")
        oComm.Parameters.Add("@dCredVigentes", SqlDbType.Xml).Value = dCredVig.DevuelveXML

        'ARCV 23-03-2010
        oComm.Parameters.Add("@cComenUF", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenUF")
        oComm.Parameters.Add("@cComenEGP", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenEGP")
        oComm.Parameters.Add("@cComenBalance", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenBalance")
        '------------------
        '-----------vvillafranca 06-08-2012-------
        oComm.Parameters.Add("@oComentarios", SqlDbType.Xml).Value = oComentarios
        '------------------


        '-----------vvillafranca 05-11-2012-------
        oComm.Parameters.Add("@bCuotaDoble", SqlDbType.Bit).Value = pbCuotaDoble
        oComm.Parameters.Add("@bCDPrimeraGratif", SqlDbType.Bit).Value = bCDPrimeraGratif
        'IIf(poMesesDobles Is Nothing, DBNull.Value, poMesesDobles.DevuelveXML)
        '------------------

        'ARCV 11-05-2011
        oComm.ExecuteNonQuery()
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
        oCal = Nothing
        oCalData = Nothing

    End Sub

    Public Sub ValidaModalidad(ByVal pnCodPers As Integer, ByVal pnCodModalidad As Integer, _
     ByRef psResp As Integer, ByRef psMensage As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ValidaModalidad", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodPersTitut", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pnModalidad", SqlDbType.Int).Value = pnCodModalidad
        oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psComentario", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        psResp = oComm.Parameters(2).Value
        psMensage = oComm.Parameters(3).Value


        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ActualizarPersonasRelacionadas(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal D As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ActualizaPersonasRelacionadas", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nRelaPers", SqlDbType.Xml).Value = D.DevuelveXML
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CredEliminaDestinosxCredito(ByVal pnCodDestino As Integer, ByVal pnCodProd As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredEliminaDestinoxCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodDestino", SqlDbType.Int).Value = pnCodDestino
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = pnCodProd
        oComm.ExecuteNonQuery()


        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function CredMantDestinosxCredito(ByVal pnCodDestino As Integer, ByVal pnCodProd As Integer) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredMantDestinosxCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodDestino", SqlDbType.Int).Value = pnCodDestino
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = pnCodProd
        oComm.Parameters.Add("@Resp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        CredMantDestinosxCredito = oComm.Parameters("@Resp").Value


        objCon = Nothing
        oComm = Nothing
    End Function
    Public Sub ActualizarCredito(ByVal D As ComTipos.clsDataset, ByVal Rela As ComTipos.clsDataset, ByVal nTipoAct As Integer, ByRef pnResNewCred As Integer, ByVal poDesemPar As ComTipos.clsDataset, ByVal pnItfValor As Double)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim nResNewCred As Integer
        Dim oFeriado As ComTipos.clsDataset
        Dim nMontoGasto As Double


        nMontoGasto = ObtieneGastoTotalCal(CDbl(D.ObtenerDato("nPrestamo")), 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))
        oFeriado = RecuperaFeriados(CDate(D.ObtenerDato("dFecDes")))



        oCal = New ComTipos.ClsCalendario
        'Se agrego instancia ya que no se generara cuota ....11/01/2010...Carlos Pua
        '***********************************
        oCalData = New ComTipos.clsDataset
        '*************************************
        'Se quito la Generacion del Calendario segun Nuevos Requerimientos de Negocios 11/01/2010.... Carlos Pua
        '**************************************************************************
        'oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(D.ObtenerDato("nCodAge")), CInt(IIf(D.ObtenerDato("nCodCred") = Nothing, "0", D.ObtenerDato("nCodCred"))), _
        '            1, CDbl(D.ObtenerDato("nPrestamo")), CInt(D.ObtenerDato("nCuotas")), CDbl(D.ObtenerDato("nTasaComp")), _
        '            CInt(D.ObtenerDato("nDiaFijo")), CDate(D.ObtenerDato("dFecDes")), CInt(D.ObtenerDato("nDiasGracia")), 0, CDbl(D.ObtenerDato("nGastos")), _
        '            CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")), IIf(CInt(D.ObtenerDato("nCodCamp")) = 0, True, False), CInt(D.ObtenerDato("nTipoPeriodo")), CInt(D.ObtenerDato("nPlazo")), pnItfValor, IIf(D.ObtenerDato("nSoloIntereses") = "1", True, False), , CDate(D.ObtenerDato("dFecPrimPago")))
        '**************************************************************************

        objCon = New ComConecta.clsConecta
        'ARCV 09-05-2011
        oComm = New SqlClient.SqlCommand("CredW_ActualizaSolcitudCredito", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("CredW_ActualizaSolcitudCredito")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
        oComm.Parameters("@nCodCred").Value = IIf(D.ObtenerDato("nCodCred") = Nothing, 0, D.ObtenerDato("nCodCred"))
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
        oComm.Parameters("@nCodAge").Value = IIf(D.ObtenerDato("nCodAge") = Nothing, 0, D.ObtenerDato("nCodAge"))
        oComm.Parameters.Add("@nProd", SqlDbType.Int)
        oComm.Parameters("@nProd").Value = D.ObtenerDato("nProd")
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int)
        oComm.Parameters("@nSubProd").Value = D.ObtenerDato("nSubProd")
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money)
        oComm.Parameters("@nPrestamo").Value = D.ObtenerDato("nPrestamo")
        oComm.Parameters.Add("@nCuotas", SqlDbType.Int)
        oComm.Parameters("@nCuotas").Value = D.ObtenerDato("nCuotas")
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int)
        oComm.Parameters("@nPlazo").Value = D.ObtenerDato("nPlazo")
        oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int)
        oComm.Parameters("@nDiaFijo").Value = D.ObtenerDato("nDiaFijo")
        oComm.Parameters.Add("@nGrupo", SqlDbType.Int)
        oComm.Parameters("@nGrupo").Value = D.ObtenerDato("nGrupo")
        oComm.Parameters.Add("@nCiclo", SqlDbType.Int)
        oComm.Parameters("@nCiclo").Value = D.ObtenerDato("nCiclo")
        oComm.Parameters.Add("@nCodLinea", SqlDbType.Int)
        oComm.Parameters("@nCodLinea").Value = D.ObtenerDato("nCodLinea")
        oComm.Parameters.Add("@nDestino", SqlDbType.Int)
        oComm.Parameters("@nDestino").Value = D.ObtenerDato("nDestino")
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int)
        oComm.Parameters("@nCodPers").Value = D.ObtenerDato("nCodPers")
        oComm.Parameters.Add("@nTasaComp", SqlDbType.Money)
        oComm.Parameters("@nTasaComp").Value = D.ObtenerDato("nTasaComp")
        oComm.Parameters.Add("@nTasaMor", SqlDbType.Money)
        oComm.Parameters("@nTasaMor").Value = 0
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money)
        oComm.Parameters("@nMontoCuota").Value = D.ObtenerDato("nMontoCuota")
        oComm.Parameters.Add("@nCondicion", SqlDbType.Int)
        oComm.Parameters("@nCondicion").Value = D.ObtenerDato("nCodModalidad")
        oComm.Parameters.Add("@nRelaPers", SqlDbType.Xml)
        oComm.Parameters("@nRelaPers").Value = Rela.DevuelveXML()
        oComm.Parameters.Add("@nTipoAct", SqlDbType.Int)
        oComm.Parameters("@nTipoAct").Value = nTipoAct
        oComm.Parameters.Add("@nGastos", SqlDbType.Money)
        oComm.Parameters("@nGastos").Value = D.ObtenerDato("nGastos")
        oComm.Parameters.Add("@dFecDes", SqlDbType.DateTime)
        oComm.Parameters("@dFecDes").Value = D.ObtenerDato("dFecDes")
        oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int)
        oComm.Parameters("@nDiasGracia").Value = CInt(D.ObtenerDato("nDiasGracia"))
        oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int)
        oComm.Parameters("@nTipoGracia").Value = CInt(D.ObtenerDato("TipoGracia"))
        oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
        oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML
        oComm.Parameters.Add("@nRespCreditoNew", SqlDbType.Int)
        oComm.Parameters("@nRespCreditoNew").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodCamp", SqlDbType.Int)
        oComm.Parameters("@nCodCamp").Value = CInt(D.ObtenerDato("nCodCamp"))
        oComm.Parameters.Add("@cCodAct", SqlDbType.VarChar)
        oComm.Parameters("@cCodAct").Value = D.ObtenerDato("cCodAct")
        oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt)
        oComm.Parameters("@nTipoPeriodo").Value = D.ObtenerDato("nTipoPeriodo")
        oComm.Parameters.Add("@nMoneda", SqlDbType.SmallInt)
        oComm.Parameters("@nMoneda").Value = D.ObtenerDato("nMoneda")
        oComm.Parameters.Add("@nOficina", SqlDbType.Int).Value = D.ObtenerDato("nOficina")
        oComm.Parameters.Add("@bTasaPref", SqlDbType.VarChar, 20).Value = D.ObtenerDato("bTasaPref")

        oComm.Parameters.Add("@Cal_Des_Par", SqlDbType.Xml)
        oComm.Parameters("@Cal_Des_Par").Value = poDesemPar.DevuelveXML
        oComm.Parameters.Add("@nTipoDesemb", SqlDbType.Int).Value = IIf(D.ObtenerDato("nTipoDesemb") = "@#err", 1, D.ObtenerDato("nTipoDesemb"))
        oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.Int).Value = IIf(D.ObtenerDato("nSoloIntereses") = "@#err", 0, D.ObtenerDato("nSoloIntereses"))
        oComm.Parameters.Add("@nTipoSegDesg", SqlDbType.Int).Value = IIf(D.ObtenerDato("nTipoSegDesg") = "@#err", 0, D.ObtenerDato("nTipoSegDesg"))

        oComm.Parameters.Add("@nPorCapTrab", SqlDbType.Money).Value = D.ObtenerDato("nPorCapTrab")
        oComm.Parameters.Add("@nPorActFijo", SqlDbType.Money).Value = D.ObtenerDato("nPorActFijo")
        oComm.Parameters.Add("@nPorCompDeuda", SqlDbType.Money).Value = D.ObtenerDato("nPorCompDeuda")
        oComm.Parameters.Add("@dFecPrimPago", SqlDbType.DateTime).Value = D.ObtenerDato("dFecPrimPago")

        oComm.ExecuteNonQuery()
        'ARCV 09-05-2011
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        nResNewCred = oComm.Parameters("@nRespCreditoNew").Value

        pnResNewCred = nResNewCred

        objCon = Nothing
        oComm = Nothing
        oCal = Nothing
        oCalData = Nothing

    End Sub

    Public Function RecuperaTasaComp(ByVal pnCodLinea As Integer) As ComTipos.clsDataset


        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredW_RecuperaTasaComp  " & pnCodLinea)

        RecuperaTasaComp = C
        C = Nothing
    End Function



    Public Sub BuscarDetalleCuotas(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByRef oPag As ComTipos.clsCredPagoCuotas)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DetalleCuotas '" & sCodAge & "','" & sCodExpe & "'")
        oPag.Cuotas = oData
        oData = Nothing
    End Sub

    Public Sub BuscarCrediPagos(ByVal dFecha As Date, _
                                ByRef oPag As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaCrediPagos '" & dFecha & "'")
        oPag = oData
        oData = Nothing
    End Sub
    Public Sub RecuperaListaCrediPago(ByVal dFecha As Date, _
                                     ByRef oPag As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaListaCrediPago '" & dFecha & "'", , , 0)
        oPag = oData
        oData = Nothing
    End Sub

    Public Sub BuscarPagosBN(ByVal dFecha As Date, _
                                ByRef oPag As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaPagosBN '" & dFecha & "'")
        oPag = oData
        oData = Nothing
    End Sub

    Public Sub RecuperaListaPagosBN(ByVal dFecha As Date, _
                                    ByRef oPag As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaListaPagosBN '" & dFecha & "'", , , 0)
        oPag = oData
        oData = Nothing
    End Sub

    Public Sub CrediPagoGrupoError(ByRef oCred As ComTipos.clsDataset, _
          ByVal dFecha As Date, _
          ByVal nCodBanco As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_WCrediPagoGrupoError '" & dFecha & "'," & nCodBanco)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub BuscarImpresion(ByVal sCodExpe As String, _
                                ByVal sCodAge As String, _
                                ByRef oDatos As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset

        oData.CargaTabla(, , "CredW_BuscarImpresion '" & sCodAge & "','" & sCodExpe & "'")
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub BuscarImpresionAVALES(ByVal sCodExpe As String, _
                                ByVal sCodAge As String, _
                                ByRef oDatos As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset

        oData.CargaTabla(, , "CredW_RecuperaAvalesCredito '" & sCodExpe & "','" & sCodAge & "'")
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub RecuperaAgencias(ByRef oAge As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_Recupera_Agencias")
        oAge = oData
        oData = Nothing
    End Sub

    Public Sub BuscaCred_x_Estado(ByRef oCred As ComTipos.clsDataset, _
                                        ByVal nCodAge As Integer, _
                                        ByVal nCodCred As Integer, _
                                        ByVal nEstado As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_BuscaCred_x_Estado " & nCodAge & "," & nCodCred & "," & nEstado)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ListarCredRefinanciar(ByRef oCred As ComTipos.clsDataset, _
                                    ByVal nCodAge As Integer, _
                                    ByVal nCodCred As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ListarCredRefinanciar " & nCodAge & "," & nCodCred)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Refinanciar(ByVal oCred As ComTipos.clsDataset, _
                                ByVal nCodAge As Integer, _
                                ByVal nCodCred As Integer)

        Dim sListaRef As String = oCred.DevuelveXML
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_Refinanciar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = sListaRef
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ListarRefinanciamientos(ByRef oCred As ComTipos.clsDataset, _
                                ByVal dFecha As Date)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ListarRefinanc '" & dFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub BuscarCredRefinanciados(ByRef oCred As ComTipos.clsDataset, _
                                ByVal nCodAge As Integer, _
                                    ByVal nCodCred As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_BuscarCredRefinanc " & nCodCred & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ExtornarRefinanc(ByVal nCodAge As Integer, _
                                    ByVal nCodCred As Integer, _
                                    ByVal dFecha As Date)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CredW_ExtornoRefinanc", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CalificacionCartera(ByVal dFecha As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CalificacionCartera", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.CommandTimeout = 600
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub TransferenciaCartera(ByVal nCodAnaAnt As Integer, _
                                    ByVal nCodAnaNue As Integer, _
                                    ByVal dFecha As Date, _
                                    ByVal oCreditos As ComTipos.clsDataset, _
                                    ByVal pnCodAgeNuevo As Integer, ByVal psCodUsu As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim sCreditos As String = oCreditos.DevuelveXML

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_TransferenciaCartera", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@nCodAnaAnt", SqlDbType.Int).Value = nCodAnaAnt
        oComm.Parameters.Add("@nCodAnaNue", SqlDbType.Int).Value = nCodAnaNue
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@oCreditos", SqlDbType.Xml).Value = sCreditos
        oComm.Parameters.Add("@nCodAgeNuevo", SqlDbType.Int).Value = pnCodAgeNuevo
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = psCodUsu

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Sub TransferenciaCarteraAGestor(ByVal nCodAnaAnt As Integer, _
                                     ByVal nCodAnaNue As Integer, _
                                     ByVal dFecha As Date, _
                                     ByVal oCreditos As ComTipos.clsDataset, _
                                     ByVal pnCodAgeNuevo As Integer, ByVal psCodUsu As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim sCreditos As String = oCreditos.DevuelveXML

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_TransferenciaCarteraAGestor", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAnaAnt", SqlDbType.Int).Value = nCodAnaAnt
        oComm.Parameters.Add("@nCodAnaNue", SqlDbType.Int).Value = nCodAnaNue
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@oCreditos", SqlDbType.Xml).Value = sCreditos
        oComm.Parameters.Add("@nCodAgeNuevo", SqlDbType.Int).Value = pnCodAgeNuevo
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = psCodUsu

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing

    End Sub


    Public Sub Anexo5_Datos(ByRef oCred As ComTipos.clsDataset, _
                        ByVal pnTipoCambio As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Anexo5_Datos " & pnTipoCambio)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ReporteFOGAPI(ByRef oCred As ComTipos.clsDataset, _
                                ByVal dFecha As Date)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ReporteFOGAPI '" & dFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ReporteFOGAPICF(ByRef oCred As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ReporteFOGAPI_CF")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ListarCred_x_Analista(ByRef oCred As ComTipos.clsDataset, _
                                ByVal nCodAnalista As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ListarCred_x_Analista " & nCodAnalista)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ListarCred_x_Analista_ParaGestor(ByRef oCred As ComTipos.clsDataset, _
                            ByVal nCodAnalista As Integer, ByVal pnCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ListarCred_x_Analista_ParaGestor " & nCodAnalista & "," & pnCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ListarGestores_x_Agencia(ByRef oCred As ComTipos.clsDataset, _
                        ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ListarGestores_x_Agencia " & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ListarAnalistas_x_Agencia(ByRef oCred As ComTipos.clsDataset, _
                            ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ListarAnalistas_x_Agencia " & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ReporteINFORCORPMora(ByRef oCred As ComTipos.clsDataset, _
                        ByVal dFecha As Date, _
                        ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ReporteINFORCORPMora '" & dFecha & "'," & nCodAge, , , 0)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ReporteINFORCORPDia(ByRef oCred As ComTipos.clsDataset, _
                                    ByVal dFecha As Date, _
                                    ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_ReporteINFORCORPDia '" & dFecha & "'," & nCodAge, , , 0)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub CargaBCPArchivo(ByVal oReg As ComTipos.clsDataset, _
                                ByVal dFecha As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim sReg As String = oReg.DevuelveXML

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CargaBCPArchivo", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@oReg", SqlDbType.Xml).Value = sReg
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Sub BCPCruceInfo(ByRef oCred As ComTipos.clsDataset, _
                        ByVal dFecha As Date)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BCPCruceInfo '" & dFecha & "'")
        oCred = oData
        oData = Nothing

    End Sub

    Public Sub BCPCruceInfo_Detalle(ByRef oCred As ComTipos.clsDataset, _
                                    ByVal dFecha As Date, _
                                    ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BCPCruceInfo_Detalle '" & dFecha & "'," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Function ObtieneVarNegocio(ByVal nCodigo As Integer) As String
        Dim oData As New ComTipos.clsDataset
        ObtieneVarNegocio = ""
        oData.CargaTabla(, , "Cred_ObtieneVarNegocio " & nCodigo)
        oData.Inicio()
        If oData.nNumFilas > 0 Then ObtieneVarNegocio = oData.ObtenerDato("cValorVar")
        oData = Nothing
    End Function

    Public Sub AsientoContable(ByVal dFecha As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_AsientoContable", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.CommandTimeout = 999999
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperaAsientoCont(ByRef oCred As ComTipos.clsDataset, _
                    ByVal dFecha As Date, _
                    ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaAsientoCont '" & dFecha & "'," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ValidaAsientoCont(ByRef oVal As ComTipos.clsDataset, _
                    ByVal dFecha As Date, _
                    ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ValidaAsientoCont '" & dFecha & "'," & nCodAge)
        oVal = oData
        oData = Nothing
    End Sub

    Public Sub CierreDiario(ByVal dFecha As Date)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CierreDiario", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.CommandTimeout = 0
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub InicioDia()
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_InicioDia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub AsientoContableDia()
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_AsientoContableDia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperarGrupos(ByVal nCodAge As Integer, _
                            ByRef oDatos As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaGrupos " & nCodAge)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub BuscarImpresionGrupal(ByRef oCred As ComTipos.clsDataset, _
                                ByVal nGrupo As Integer, _
                                ByVal nCiclo As Integer, _
                                ByVal nCodAge As Integer, _
                                Optional ByVal nEstado As Integer = -1)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_BuscarImpresionGrupal " & nGrupo & "," & nCiclo & "," & nCodAge & "," & nEstado)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub BuscarDesembolsoGrupal(ByRef oCred As ComTipos.clsDataset, _
                            ByVal nGrupo As Integer, _
                            ByVal nCiclo As Integer, _
                            ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_BuscarDesembolsoGrupal " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub CargaDatosDesembolsoGrupo(ByRef oCred As ComTipos.clsDataset, _
                        ByVal nGrupo As Integer, _
                        ByVal nCiclo As Integer, _
                        ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_CargaDatosDesembolsoGrupo " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub


    Public Sub CalendarioGrupo(ByRef oCred As ComTipos.clsDataset, _
                    ByVal nGrupo As Integer, _
                    ByVal nCiclo As Integer, _
                    ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_CalendarioGrupo " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub CargaInterfazDesembGrupo(ByVal nCodAge As Integer, _
                ByRef oDatos As ComTipos.clsDataset, _
                ByRef oBancos As ComTipos.clsDataset)
        oCons = New ComConstante.clsConstante
        Call RecuperarGrupos(nCodAge, oDatos)
        oBancos = oCons.RecuperaConstantes(2000)
        oCons = Nothing
    End Sub

    Public Sub RecuperaRelaCred(ByRef oCred As ComTipos.clsDataset, _
                ByVal nCodCred As Integer, _
                ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaRelaCred " & nCodCred & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Function RecuperaCredVig(ByVal pCodPers As String) As Integer
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaNumCredVig " & pCodPers)
        RecuperaCredVig = oData.ObtenerDato("nNumCredVig")
        oData = Nothing

    End Function

    Public Function RecuperaCredAvalados(ByVal pCodPers As String) As Integer
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaNumCredAvalados " & pCodPers)
        RecuperaCredAvalados = oData.ObtenerDato("nNumCredAval")
        oData = Nothing
    End Function

    Public Function RecuperaCredExcepciones(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaCredExcepcion " & pnCodPers)
        RecuperaCredExcepciones = oData
        oData = Nothing
    End Function
    Public Function BuscaExpediente_Grupo(ByVal nGrupo As Integer, _
                                    ByVal nCiclo As Integer, _
                                    ByVal dFecha As Date) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscaExpediente_Grupo " & nGrupo & "," & nCiclo & ",'" & dFecha & "'")
        Return oData
        'oData = Nothing
    End Function

    Public Sub GrabarPagoCuotasMicroSeguro(ByVal dFecha As Date, ByVal cPoliza As String, ByVal nCodAge As Integer, _
                            ByVal nMonto As Double, ByVal sUser As String, ByRef nNroTrans As Integer, ByVal cSolicitud As String, ByVal nMes As Integer, ByVal nAño As Integer, ByRef nValida As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GrabaPagoMicroSeguro", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nPoliza", SqlDbType.VarChar, 8).Value = cPoliza
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = sUser
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nSolicitud", SqlDbType.VarChar, 8).Value = cSolicitud
        oComm.Parameters.Add("@nMes", SqlDbType.Int).Value = nMes
        oComm.Parameters.Add("@nAño", SqlDbType.Int).Value = nAño
        oComm.Parameters.Add("@nValida", SqlDbType.Int)
        oComm.Parameters.Item("@nValida").Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        nNroTrans = oComm.Parameters.Item("@nONROTRANS").Value
        nValida = oComm.Parameters.Item("@nValida").Value

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarPagoCuotas_WIN(ByVal sCodExpe As String, _
                              ByVal sCodAge As String, _
                              ByVal sCodAgeTrans As String, _
                              ByVal dFecha As Date, _
                              ByVal nMonto As Double, _
                              ByVal nCodOpe As Integer, _
                              ByVal nCodBanco As Integer, _
                              ByVal nITF As Double, _
                          ByRef nOCAPITAL As Double, _
                          ByRef nOINTERES As Double, _
                          ByRef nOGASTOS As Double, _
                          ByRef nOMORA As Double, _
                          ByRef nOIGV As Double, _
                          ByRef nOSALDO As Double, _
                          ByRef nOPROXCUOTA As Integer, _
                          ByRef nONROTRANS As Integer, _
                          ByVal sUser As String, _
                          ByRef sOAGEDIR As String, _
                          ByRef sOAGETEL As String, _
                          ByRef sODESCOPE As String, _
                          ByRef sPCOMENTB As String, _
                          ByRef sSCOMENTB As String, _
                          Optional ByVal sNumCheque As String = "", _
                          Optional ByRef sProxVenc As String = "", _
                          Optional ByVal nPagoCta As Integer = 0, _
                          Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        'ARCV 11-05-2011
        oComm = New SqlClient.SqlCommand("Cred_PagoCuotas", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("Cred_PagoCuotas")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output
        '---------------

        oComm.ExecuteNonQuery()
        'ARCV 11-05-2011
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarPagoCuotas_WIN_Horizontal(ByVal sCodExpe As String, _
                          ByVal sCodAge As String, _
                          ByVal sCodAgeTrans As String, _
                          ByVal dFecha As Date, _
                          ByVal nMonto As Double, _
                          ByVal nCodOpe As Integer, _
                          ByVal nCodBanco As Integer, _
                          ByVal nITF As Double, _
                      ByRef nOCAPITAL As Double, _
                      ByRef nOINTERES As Double, _
                      ByRef nOGASTOS As Double, _
                      ByRef nOMORA As Double, _
                      ByRef nOIGV As Double, _
                      ByRef nOSALDO As Double, _
                      ByRef nOPROXCUOTA As Integer, _
                      ByRef nONROTRANS As Integer, _
                      ByVal sUser As String, _
                      ByRef sOAGEDIR As String, _
                      ByRef sOAGETEL As String, _
                      ByRef sODESCOPE As String, _
                      ByRef sPCOMENTB As String, _
                      ByRef sSCOMENTB As String, _
                      Optional ByVal sNumCheque As String = "", _
                      Optional ByRef sProxVenc As String = "", _
                      Optional ByVal nPagoCta As Integer = 0, _
                      Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PagoCuotas_Horizontal", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output
        '---------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        objCon = Nothing
        oComm = Nothing
    End Sub


    Public Sub GrabarPagoCuotas_WIN_Verical(ByVal sCodExpe As String, _
                              ByVal sCodAge As String, _
                              ByVal sCodAgeTrans As String, _
                              ByVal dFecha As Date, _
                              ByVal nMonto As Double, _
                              ByVal nCodOpe As Integer, _
                              ByVal nCodBanco As Integer, _
                              ByVal nITF As Double, _
                          ByRef nOCAPITAL As Double, _
                          ByRef nOINTERES As Double, _
                          ByRef nOGASTOS As Double, _
                          ByRef nOMORA As Double, _
                          ByRef nOIGV As Double, _
                          ByRef nOSALDO As Double, _
                          ByRef nOPROXCUOTA As Integer, _
                          ByRef nONROTRANS As Integer, _
                          ByVal sUser As String, _
                          ByRef sOAGEDIR As String, _
                          ByRef sOAGETEL As String, _
                          ByRef sODESCOPE As String, _
                          ByRef sPCOMENTB As String, _
                          ByRef sSCOMENTB As String, _
                          Optional ByVal sNumCheque As String = "", _
                          Optional ByRef sProxVenc As String = "", _
                          Optional ByVal nPagoCta As Integer = 0, Optional ByVal poCuotas As ComTipos.clsDataset = Nothing, _
                          Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PagoCuotas_Vertical", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque
        oComm.Parameters.Add("@oCalenAmortizado", SqlDbType.Xml).Value = poCuotas.DevuelveXML

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output
        '---------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        objCon = Nothing
        oComm = Nothing
    End Sub


    Public Sub GrabarNotaCreditoPagoCuotas_WIN_Verical(ByVal sCodExpe As String, _
                              ByVal sCodAge As String, _
                              ByVal sCodAgeTrans As String, _
                              ByVal dFecha As Date, _
                              ByVal nMonto As Double, _
                              ByVal nCodOpe As Integer, _
                              ByVal nCodBanco As Integer, _
                              ByVal nITF As Double, _
                          ByRef nOCAPITAL As Double, _
                          ByRef nOINTERES As Double, _
                          ByRef nOGASTOS As Double, _
                          ByRef nOMORA As Double, _
                          ByRef nOIGV As Double, _
                          ByRef nOSALDO As Double, _
                          ByRef nOPROXCUOTA As Integer, _
                          ByRef nONROTRANS As Integer, _
                          ByVal sUser As String, _
                          ByRef sOAGEDIR As String, _
                          ByRef sOAGETEL As String, _
                          ByRef sODESCOPE As String, _
                          ByRef sPCOMENTB As String, _
                          ByRef sSCOMENTB As String, _
                          Optional ByVal sNumCheque As String = "", _
                          Optional ByRef sProxVenc As String = "", _
                          Optional ByVal nPagoCta As Integer = 0, Optional ByVal poCuotas As ComTipos.clsDataset = Nothing)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PagoCuotas_Condonacion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque
        oComm.Parameters.Add("@oCalenAmortizado", SqlDbType.Xml).Value = poCuotas.DevuelveXML

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarPagoCuotasCastigados_WIN(ByVal sCodExpe As String, _
                            ByVal sCodAge As String, _
                            ByVal sCodAgeTrans As String, _
                            ByVal dFecha As Date, _
                            ByVal nMonto As Double, _
                            ByVal nCodOpe As Integer, _
                            ByVal nCodBanco As Integer, _
                            ByVal nITF As Double, _
                        ByRef nOCAPITAL As Double, _
                        ByRef nOINTERES As Double, _
                        ByRef nOGASTOS As Double, _
                        ByRef nOMORA As Double, _
                        ByRef nOIGV As Double, _
                        ByRef nOSALDO As Double, _
                        ByRef nOPROXCUOTA As Integer, _
                        ByRef nONROTRANS As Integer, _
                        ByVal sUser As String, _
                          ByRef sOAGEDIR As String, _
                          ByRef sOAGETEL As String, _
                          ByVal nPagoCta As Integer, _
                          ByRef sODESCOPE As String, _
                          ByRef nOSEGURO As Double, _
                          ByVal sNumCheque As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PagoCuotasCastigados", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----

        oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta

        'ARCV 29-09-2009
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque
        '---------------
        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----

        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value

        'ARCV 11-12-2009
        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value
        '----------------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarPagoCuotasCastigados_FIDEICOMISO(ByVal sCodExpe As String, _
                            ByVal sCodAge As String, _
                            ByVal sCodAgeTrans As String, _
                            ByVal dFecha As Date, _
                            ByVal nMonto As Double, _
                            ByVal nCodOpe As Integer, _
                            ByVal nCodBanco As Integer, _
                            ByVal nITF As Double, _
                        ByRef nOCAPITAL As Double, _
                        ByRef nOINTERES As Double, _
                        ByRef nOGASTOS As Double, _
                        ByRef nOMORA As Double, _
                        ByRef nOIGV As Double, _
                        ByRef nOSALDO As Double, _
                        ByRef nOPROXCUOTA As Integer, _
                        ByRef nONROTRANS As Integer, _
                        ByVal sUser As String, _
                          ByRef sOAGEDIR As String, _
                          ByRef sOAGETEL As String, _
                          ByVal nPagoCta As Integer, _
                          ByRef sODESCOPE As String, _
                          ByRef nOSEGURO As Double, _
                          ByVal sNumCheque As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("DBCORE..Cred_PagoCuotasCastigados", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----

        oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta

        'ARCV 29-09-2009
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque
        '---------------
        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----

        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value

        'ARCV 11-12-2009
        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value
        '----------------

        objCon = Nothing
        oComm = Nothing
    End Sub


    Public Function ObtieneGastoTotalCalNew(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Double

        Dim R As ComTipos.clsDataset

        R = New ComTipos.clsDataset
        R.CargaTabla(, , "Cred_RecuperaGastosALDesemb " & pnCodAge & "," & pnCodCred)

        ObtieneGastoTotalCalNew = CDbl(R.ObtenerDato("nValorGasto"))

        R = Nothing

    End Function

    Public Sub DesembolsarCredito_WIN(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByVal sCodAgeTrans As String, _
                                    ByVal dFecha As Date, _
                                    ByVal nPrestamo As Double, _
                                    ByVal nCodOpe As Integer, _
                                    ByVal nCodBanco As Integer, _
                                    ByVal nITF As Double, _
                                    ByVal nCuotas As Integer, _
                                    ByVal nTasa As Double, _
                                    ByVal nDia As Integer, _
                                    ByVal nDiasGracia As Integer, _
                                    ByVal nTipoGracia As Integer, _
                                    ByVal nGastos As Double, _
                                ByRef dOFecVenc As Date, _
                                ByRef cOPagare As String, _
                                ByRef cODocEmi As String, _
                                ByRef cODomEmi As String, _
                                ByRef cOConyuge As String, _
                                ByRef cODocCon As String, _
                                ByRef cODomCon As String, _
                            ByRef cODpto As String, _
                            ByRef cOProv As String, _
                            ByRef cODist As String, _
                            ByRef cOLinea As String, _
                            ByRef cOMoneda As String, _
                            ByRef bODNI As Boolean, _
                            ByVal sUser As String, _
                            ByRef cOCodBCP As String, _
                            ByRef cOCodBN As String, _
                            ByVal nProd As Integer, _
                            ByVal nSubProd As Integer, _
                            ByRef cOGrupo As String, _
                            ByVal nPlazo As Integer, _
                            ByVal nTipoPeriodo As Integer, _
                            ByRef sPCOMENTC As String, _
                              ByRef sSCOMENTC As String, _
                              ByRef sCodAgeBN As Integer, ByRef nNroTrans As Integer, _
                              Optional ByVal sNumCheque As String = "", _
                              Optional ByVal sRUC As String = "", _
                              Optional ByRef sZona As String = "", Optional ByVal nValorITF As Double = 0, Optional ByVal pnSoloIntereses As Integer = 0, _
                              Optional ByVal nAgenciaBNOfComp As Integer = 0, _
                              Optional ByVal sClaveBNOfComp As String = "")

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim sGastos As String
        Dim sCalend As String
        Dim oTemp As Object
        'Dim MatGastosCol As Object
        Dim MatCalendCol As Object
        Dim dsCalend As New ComTipos.clsDataset
        Dim dsGastos As New ComTipos.clsDataset
        Dim dsCamp As New ComTipos.clsDataset
        Dim dsCredDesPar As New ComTipos.clsDataset
        Dim dsCalenDesPar As New ComTipos.clsDataset
        Dim dsCalDesembRenov As New ComTipos.clsDataset
        Dim oCalend As New ComTipos.ClsCalendario
        Dim i As Integer
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset
        Dim MatCalenCuotasDesemParDetalle As ComTipos.clsDataset
        Dim nPrestamoTempo As Double
        Dim nTasaDesgravamen As Double
        Dim pbSoloIntereses As Boolean

        Dim dFecPrimPago As Date 'ARCV 20-01-2010
        '---------- vvillafranca 05-11-2012-
        Dim bCDPrimeraGratif As Boolean
        Dim pbCuotaDoble As Boolean
        'vvillafranca 09-11-2012 (valida si permite redoondeo)
        Dim bRedondeo As Boolean
        Dim oRedondeo As ComTipos.clsDataset
        oRedondeo = New ComTipos.clsDataset
        oRedondeo.CargaTabla(, , "Cred_RecuperaRedondeoxAgencia " & sCodAge)
        bRedondeo = CBool(oRedondeo.ObtenerDato("bcuotasindec"))
  
        '-----------------
        dsCamp.CargaTabla(, , "CredW_RecuperaCreditoDatos " & sCodExpe & "," & sCodAge) 'vvillafranca aumente bcuota doble

        pbSoloIntereses = IIf(dsCamp.ObtenerDato("nSoloIntereses") <> 0, True, False)  'ARCV 10-09-2009

        dFecPrimPago = CDate(dsCamp.ObtenerDato("dFecPrimPago")) 'ARCV 20-01-2010
        pbCuotaDoble = CBool(dsCamp.ObtenerDato("bCuotaDoble")) 'vvillafranca 05-11-2012
        bCDPrimeraGratif = CBool(dsCamp.ObtenerDato("bCDPrimeraGratif")) 'vvillafranca 05-11-2012

        dsCredDesPar.CargaTabla(, , "Cred_RecuperaDatosDesemParcial " & sCodAge & "," & sCodExpe)
        dsCalenDesPar.CargaTabla(, , "Cred_RecuperaCalDesParcial " & sCodAge & "," & sCodExpe)




        MatCalenCuotasDesemParDetalle = New ComTipos.clsDataset


        'ARCV 21-05-2009
        nMontoGasto = ObtieneGastoTotalCalNew(CInt(sCodAge), CInt(sCodExpe))
        oFeriado = RecuperaFeriados(dFecha)

        'nMontoGasto = ObtieneGastoTotalCal(nPrestamo, 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))
        nTasaDesgravamen = RecuperaTasaSeguroDesgravamen(CInt(sCodExpe), CInt(sCodAge), _
                                    nPrestamo, 1)
        '----------
        'MatCalenCuotasDesemParDetalle = New ComTipos.clsDataset
        'MatCalenCuotasDesemParDetalle = Nothing

        If dsCredDesPar.ObtenerDato("nTipoDesemb") = "2" Then
            nPrestamoTempo = 0
            For i = dsCalenDesPar.nNumFilas To 1 Step -1
                If i > CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
                    dsCalenDesPar.EliminaRegistro(i)
                Else
                    If CInt(dsCalenDesPar.ObtenerDato("nItem", i)) = CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
                        dsCalenDesPar.AsignarDato(Format(dFecha, "dd/MM/yyyy"), "dFecha", i)
                    End If
                    nPrestamoTempo = nPrestamoTempo + CDbl(dsCalenDesPar.ObtenerDato("nPrestamo", i))
                End If
            Next

            'oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamoTempo, nCuotas, nTasa, nDia, dsCredDesPar.ObtenerDato("dFecVig"), nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, False, nTasaDesgravamen, True, dsCalenDesPar, MatCalenCuotasDesemParDetalle, _
            '                                dFecPrimPago) 'ARCV 20-01-2010
            oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamoTempo, nCuotas, nTasa, nDia, dsCredDesPar.ObtenerDato("dFecVig"), nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, False, nTasaDesgravamen, True, dsCalenDesPar, MatCalenCuotasDesemParDetalle, _
                                          dFecPrimPago, , pbCuotaDoble, bCDPrimeraGratif, bRedondeo) 'vvillafranca 05-11-2012
            'dsCalDesembRenov = RenovarDesembolso(sCodAge, sCodExpe, dFecha, nValorITF, nPrestamo, pnSoloIntereses)
        Else
            'Napoleon Silva - Campañas Enlatados 12/09/2011
            'oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamo, nCuotas, nTasa, nDia, dFecha, nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, pbSoloIntereses, nTasaDesgravamen, , , , _
            '                                dFecPrimPago, IIf(CInt(dsCamp.ObtenerDato("nEnlatado")) > 0, CDbl(dsCamp.ObtenerDato("nMontoCuota")), -1))
            oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamo, nCuotas, nTasa, nDia, dFecha, nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, pbSoloIntereses, nTasaDesgravamen, , , , _
                                                      dFecPrimPago, IIf(CInt(dsCamp.ObtenerDato("nEnlatado")) > 0, CDbl(dsCamp.ObtenerDato("nMontoCuota")), -1), pbCuotaDoble, bCDPrimeraGratif, bRedondeo)

        End If

        dsCamp = Nothing

        oCalend = Nothing

        Call dsGastos.CargaTabla(, , "CredW_RecuperaGastosAlDesembDet " & nPrestamo & "," & CInt(sCodExpe) & "," & CInt(sCodAge))

        'VVILLAFRANCA

        '------------------

        sGastos = dsGastos.DevuelveXML

        dsGastos = Nothing

        'If dsCredDesPar.ObtenerDato("nTipoDesemb") <> "2" Then

        ReDim MatCalendCol(7)
        MatCalendCol(1) = "nNroCuota"
        MatCalendCol(2) = "dFecVenc"
        MatCalendCol(3) = "nCapital"
        MatCalendCol(4) = "nIntComp"
        MatCalendCol(5) = "nIntMor"
        MatCalendCol(6) = "nIGV"
        MatCalendCol(7) = "nIntGra" 'ARCV 30-10-2007

        With dsCalend
            .CrearVacio(MatCalendCol)
            For i = 1 To UBound(oTemp)
                .AddRegistro()
                .Seek(i)
                .AsignarDato(oTemp(i, 2), "nNroCuota")
                .AsignarDato(CDate(oTemp(i, 1)).ToString("yyyyMMdd"), "dFecVenc")
                .AsignarDato(oTemp(i, 4), "nCapital")
                .AsignarDato(oTemp(i, 5), "nIntComp")
                .AsignarDato(0, "nIntMor")
                .AsignarDato(oTemp(i, 6), "nIGV")
                .AsignarDato(oTemp(i, 9), "nIntGra") 'ARCV 30-10-2007
            Next i
        End With

        sCalend = dsCalend.DevuelveXML
        dOFecVenc = CDate(oTemp(UBound(oTemp), 1)) 'Ultima fecha del Calendario

        'Else
        'sCalend = dsCalDesembRenov.DevuelveXML
        'dOFecVenc = CDate(dsCalDesembRenov.ObtenerDato("dFecVenc", dsCalDesembRenov.nNumFilas)) 'Ultima fecha del Calendario
        'End If


        dsCalend = Nothing

        objCon = New ComConecta.clsConecta
        'ARCV 11-05-2011
        oComm = New SqlClient.SqlCommand("Cred_DesembolsarCredito", objCon.DameConexion)

        'oComm = New SqlClient.SqlCommand("Cred_DesembolsarCredito")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        'oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco

        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF
        'Se Agrego Parametro AgenciaBN...Carlos Pua 03/12/09
        oComm.Parameters.Add("@nCodAgenciaBN", SqlDbType.Int).Value = sCodAgeBN
        '---------------------------------------------------------------------
        oComm.Parameters.Add("@XCalend", SqlDbType.Xml).Value = sCalend
        oComm.Parameters.Add("@XGastos", SqlDbType.Xml).Value = sGastos

        oComm.Parameters.Add("@cOPagare", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOPagare").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODocEmi", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODocEmi").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODomEmi", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cODomEmi").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOConyuge", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOConyuge").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODocCon", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODocCon").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODomCon", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cODomCon").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cODpto", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODpto").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOProv", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOProv").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODist", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODist").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOLinea", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOLinea").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOMoneda", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOMoneda").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@bODNI", SqlDbType.Bit)
        oComm.Parameters.Item("@bODNI").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 31-10-2007
        oComm.Parameters.Add("@cOCodBCP", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOCodBCP").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOCodBN", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOCodBN").Direction = ParameterDirection.Output
        '-------------

        'ARCV 02-11-2007
        oComm.Parameters.Add("@cOGrupo", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOGrupo").Direction = ParameterDirection.Output
        '----------

        'ARCV 07-07-2008
        'oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cPComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque
        oComm.Parameters.Add("@cRUC", SqlDbType.VarChar, 20).Value = sRUC
        oComm.Parameters.Add("@cOZona", SqlDbType.VarChar, 20).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@XCalenDesemPar", SqlDbType.Xml).Value = MatCalenCuotasDesemParDetalle.DevuelveXML
        'mestrada
        oComm.Parameters.Add("@nNroTransaccion", SqlDbType.Int).Direction = ParameterDirection.Output

        'ARCV 17-07-2012
        oComm.Parameters.Add("@nAgenciaBNOfComp", SqlDbType.Int).Value = nAgenciaBNOfComp
        oComm.Parameters.Add("@cClaveBNOfComp", SqlDbType.VarChar, 4).Value = sClaveBNOfComp
        '---------

        oComm.ExecuteNonQuery()

        'ARCV 11-05-2011
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        cOPagare = Trim(oComm.Parameters.Item("@cOPagare").Value)
        cODocEmi = Trim(oComm.Parameters.Item("@cODocEmi").Value)
        cODomEmi = Trim(oComm.Parameters.Item("@cODomEmi").Value)
        cOConyuge = Trim(oComm.Parameters.Item("@cOConyuge").Value)
        cODocCon = Trim(oComm.Parameters.Item("@cODocCon").Value)
        cODomCon = Trim(oComm.Parameters.Item("@cODomCon").Value)

        cODpto = Trim(oComm.Parameters.Item("@cODpto").Value)
        cOProv = Trim(oComm.Parameters.Item("@cOProv").Value)
        cODist = Trim(oComm.Parameters.Item("@cODist").Value)
        cOLinea = Trim(oComm.Parameters.Item("@cOLinea").Value)
        cOMoneda = Trim(oComm.Parameters.Item("@cOMoneda").Value)
        bODNI = oComm.Parameters.Item("@bODNI").Value

        'ARCV 31-10-2007
        cOCodBCP = Trim(IIf(IsDBNull(oComm.Parameters.Item("@cOCodBCP").Value), "", oComm.Parameters.Item("@cOCodBCP").Value))
        cOCodBN = Trim(IIf(IsDBNull(oComm.Parameters.Item("@cOCodBN").Value), "", oComm.Parameters.Item("@cOCodBN").Value))
        '-------
        cOGrupo = Trim(oComm.Parameters.Item("@cOGrupo").Value) 'ARCV 02-11-2007

        'ARCV 07-07-2008
        'sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        'sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sPCOMENTC = CStr(oComm.Parameters.Item("@cPComentC").Value)
        sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        sZona = CStr(oComm.Parameters.Item("@cOZona").Value)
        'MESTRADA 19-08-2011
        nNroTrans = CInt(oComm.Parameters.Item("@nNroTransaccion").Value)
        '------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function GeneraCalandarioHojaResumen(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByVal dFecha As Date, _
                                    ByVal nPrestamo As Double, _
                                    ByVal nCuotas As Integer, _
                                    ByVal nTasa As Double, _
                                    ByVal nDia As Integer, _
                                    ByVal nDiasGracia As Integer, _
                                    ByVal nTipoGracia As Integer, _
                                    ByVal nGastos As Double, _
                            ByVal nProd As Integer, _
                            ByVal nSubProd As Integer, _
                            ByVal nPlazo As Integer, _
                            ByVal nTipoPeriodo As Integer, _
                            Optional ByVal nValorITF As Double = 0) As ComTipos.clsDataset


        Dim oTemp As Object

        Dim dsCamp As New ComTipos.clsDataset

        Dim oCalend As New ComTipos.ClsCalendario
        Dim dsCalend As New ComTipos.clsDataset
        Dim dsCredDesPar As New ComTipos.clsDataset
        Dim dsCalenDesPar As New ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset
        Dim MatCalendCol As Object
        Dim i As Integer
        Dim nTasaDesgravamen As Double
        Dim pbSoloIntereses As Boolean
        Dim nTotal As Double
        Dim nItf As Double
        Dim nPrestamoTempo As Double
        Dim dFecPrimPago As Date 'ARCV 20-01-2010
        Dim MatCalenCuotasDesemParDetalle As ComTipos.clsDataset
        Dim oMontoDesParcial As New ComTipos.clsDataset ' cpua 14/04/2010

        dsCamp.CargaTabla(, , "CredW_RecuperaCreditoDatos " & sCodExpe & "," & sCodAge)
        oMontoDesParcial.CargaTabla(, , "Cred_DevuelveDesParMontoPrestamoActual " & sCodAge & "," & sCodExpe)

        pbSoloIntereses = IIf(dsCamp.ObtenerDato("nSoloIntereses") <> 0, True, False)  'ARCV 10-09-2009

        dFecPrimPago = CDate(dsCamp.ObtenerDato("dFecPrimPago")) 'ARCV 20-01-2010


        'ARCV 21-05-2009
        nMontoGasto = ObtieneGastoTotalCalNew(CInt(sCodAge), CInt(sCodExpe))
        oFeriado = RecuperaFeriados(dFecha)

        'nMontoGasto = ObtieneGastoTotalCal(nPrestamo, 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))
        nTasaDesgravamen = RecuperaTasaSeguroDesgravamen(CInt(sCodExpe), CInt(sCodAge), _
                                    nPrestamo, 1)
        '----------
        'MatCalenCuotasDesemParDetalle = New ComTipos.clsDataset
        'MatCalenCuotasDesemParDetalle = Nothing

        'If dsCredDesPar.ObtenerDato("nTipoDesemb") = "2" Then
        '    nPrestamoTempo = 0
        '    For i = dsCalenDesPar.nNumFilas To 1 Step -1
        '        If i > CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
        '            dsCalenDesPar.EliminaRegistro(i)
        '        Else
        '            If CInt(dsCalenDesPar.ObtenerDato("nItem", i)) = CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
        '                dsCalenDesPar.AsignarDato(Format(dFecha, "dd/MM/yyyy"), "dFecha", i)
        '            End If
        '            nPrestamoTempo = nPrestamoTempo + CDbl(dsCalenDesPar.ObtenerDato("nPrestamo", i))
        '        End If
        '    Next

        '    oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamoTempo, nCuotas, nTasa, nDia, dsCredDesPar.ObtenerDato("dFecVig"), nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, False, nTasaDesgravamen, True, dsCalenDesPar, MatCalenCuotasDesemParDetalle, _
        '                                    dFecPrimPago) 'ARCV 20-01-2010
        '    'dsCalDesembRenov = RenovarDesembolso(sCodAge, sCodExpe, dFecha, nValorITF, nPrestamo, pnSoloIntereses)
        'Else
        dsCredDesPar.CargaTabla(, , "Cred_RecuperaDatosDesemParcial " & sCodAge & "," & sCodExpe)
        dsCalenDesPar.CargaTabla(, , "Cred_RecuperaCalDesParcial_HojaResumen " & sCodAge & "," & sCodExpe & ", '" & dFecha & "'")

        MatCalenCuotasDesemParDetalle = New ComTipos.clsDataset

        If dsCredDesPar.ObtenerDato("nTipoDesemb") = "2" Then
            nPrestamoTempo = 0
            'For i = dsCalenDesPar.nNumFilas To 1 Step -1
            '    If i > CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
            '        dsCalenDesPar.EliminaRegistro(i)
            '    Else
            '        If CInt(dsCalenDesPar.ObtenerDato("nItem", i)) = CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
            '            dsCalenDesPar.AsignarDato(Format(dFecha, "dd/MM/yyyy"), "dFecha", i)
            '        End If
            '        nPrestamoTempo = nPrestamoTempo + CDbl(dsCalenDesPar.ObtenerDato("nPrestamo", i))
            '    End If
            'Next
            nTasaDesgravamen = 0
            If oMontoDesParcial.nNumFilas > 0 Then
                nPrestamoTempo = oMontoDesParcial.ObtenerDato("nPrestamo")
            End If
            'oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamoTempo, nCuotas, nTasa, nDia, dsCredDesPar.ObtenerDato("dFecVig"), nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, False, nTasaDesgravamen, True, dsCalenDesPar, MatCalenCuotasDesemParDetalle, _
            '                               dFecPrimPago) 'ARCV 20-01-2010

            oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamoTempo, nCuotas, nTasa, nDia, dsCredDesPar.ObtenerDato("dFecVig"), nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, True, nTipoPeriodo, nPlazo, nValorITF, False, nTasaDesgravamen, True, dsCalenDesPar, MatCalenCuotasDesemParDetalle, _
                                            dFecPrimPago) 'ARCV 20-01-2010

        Else
            'oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamo, nCuotas, nTasa, nDia, dFecha, nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, pbSoloIntereses, nTasaDesgravamen, , , , _
            '                               dFecPrimPago)
            oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamo, nCuotas, nTasa, nDia, dFecha, nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, True, nTipoPeriodo, nPlazo, nValorITF, pbSoloIntereses, nTasaDesgravamen, , , , _
                                           dFecPrimPago)
        End If

        'End If


        ReDim MatCalendCol(11)
        MatCalendCol(1) = "nNroCuota"
        MatCalendCol(2) = "dFecVenc"
        MatCalendCol(3) = "nCapital"
        MatCalendCol(4) = "nIntComp"
        MatCalendCol(5) = "nIntMor"
        MatCalendCol(6) = "nIGV"
        MatCalendCol(7) = "nIntGra" 'ARCV 30-10-2007
        MatCalendCol(8) = "nItf"
        MatCalendCol(9) = "nGasto" 'cpua 06/04/2010
        MatCalendCol(10) = "nSeguro" 'cpua 06/04/2010
        MatCalendCol(11) = "nTotal"


        With dsCalend
            .CrearVacio(MatCalendCol)
            For i = 1 To UBound(oTemp)
                .AddRegistro()
                .Seek(i)
                .AsignarDato(oTemp(i, 2), "nNroCuota")
                .AsignarDato(CDate(oTemp(i, 1)).ToString("yyyyMMdd"), "dFecVenc")
                .AsignarDato(oTemp(i, 4), "nCapital")
                .AsignarDato(oTemp(i, 5), "nIntComp")
                .AsignarDato(0, "nIntMor")
                .AsignarDato(oTemp(i, 6), "nIGV")
                .AsignarDato(oTemp(i, 9), "nIntGra") 'ARCV 30-10-2007
                .AsignarDato(oTemp(i, 7), "nGasto") 'cpua 06/04/2010
                .AsignarDato(oTemp(i, 10), "nSeguro") 'cpua 06/04/2010
                '.AsignarDato(0, "nGasto") 'cpua 06/04/2010
                '.AsignarDato(0, "nSeguro") 'cpua 06/04/2010

                'ARCV 24-05-2010
                'nItf = ((CDbl(oTemp(i, 4)) + CDbl(oTemp(i, 5)) + CDbl(oTemp(i, 6)) + CDbl(oTemp(i, 9)) + CDbl(oTemp(i, 7)) + CDbl(oTemp(i, 10))) * 1 * nValorITF)
                'nTotal = ((CDbl(oTemp(i, 4)) + CDbl(oTemp(i, 5)) + CDbl(oTemp(i, 6)) + CDbl(oTemp(i, 9)) + CDbl(oTemp(i, 7)) + CDbl(oTemp(i, 10))) * 1 * nValorITF) + (CDbl(oTemp(i, 4)) + CDbl(oTemp(i, 5)) + CDbl(oTemp(i, 6)) + CDbl(oTemp(i, 9)) + CDbl(oTemp(i, 7)) + CDbl(oTemp(i, 10)))
                nItf = 0
                nTotal = (CDbl(oTemp(i, 4)) + CDbl(oTemp(i, 5)) + CDbl(oTemp(i, 6)) + CDbl(oTemp(i, 9)) + CDbl(oTemp(i, 7)) + CDbl(oTemp(i, 10)))
                '----------------------------------

                'nItf = CDbl(oTemp(i, 3)) * nValorITF
                'nTotal = CDbl(oTemp(i, 3))
                .AsignarDato(CDbl(nItf), "nITF")
                .AsignarDato(CDbl(nTotal), "nTotal")
                '----------
            Next i
        End With


        GeneraCalandarioHojaResumen = dsCalend
        dsCamp = Nothing
        dsCalend = Nothing
        oCalend = Nothing
    End Function

    Public Sub BuscarDesembolsoGrupal_WIN(ByRef oCred As ComTipos.clsDataset, _
                        ByVal nGrupo As Integer, _
                        ByVal nCiclo As Integer, _
                        ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarDesembolsoGrupal " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub CargaDatosDesembolsoGrupo_WIN(ByRef oCred As ComTipos.clsDataset, _
                    ByVal nGrupo As Integer, _
                    ByVal nCiclo As Integer, _
                    ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CargaDatosDesembolsoGrupo " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Garantia_Insertar(ByVal dFecReg As Date, _
                                ByVal nTipoGarant As Integer, _
                                ByVal nClaseGarant As Integer, _
                                ByVal nTipoDoc As Integer, _
                                ByVal sNroDoc As String, _
                                ByVal nCodEmis As Integer, _
                                ByVal nMoneda As Integer, _
                                ByVal sDescripcion As String, _
                                ByVal sCodZona As String, _
                                ByVal nEstado As Integer, _
                                ByVal nTasacion As Double, _
                                ByVal nRealizacion As Double, _
                                ByVal nGravament As Double, _
                                ByVal nSaldo As Double, _
                                ByVal bGarantReal As Integer, _
                                ByVal nTipoRealizacion As Integer, _
                                ByVal sComentario As String, _
                                ByVal sDireccion As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Garantia_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@nTipoGarant", SqlDbType.Int).Value = nTipoGarant
        oComm.Parameters.Add("@nClaseGarant", SqlDbType.Int).Value = nClaseGarant
        oComm.Parameters.Add("@nTipoDoc", SqlDbType.Int).Value = nTipoDoc
        oComm.Parameters.Add("@cNroDoc", SqlDbType.VarChar, 50).Value = sNroDoc
        oComm.Parameters.Add("@nCodEmis", SqlDbType.Int).Value = nCodEmis
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Value = sDireccion
        oComm.Parameters.Add("@cCodZona", SqlDbType.Char, 12).Value = sCodZona
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = nEstado
        oComm.Parameters.Add("@nTasacion", SqlDbType.Money).Value = nTasacion
        oComm.Parameters.Add("@nRealizacion", SqlDbType.Money).Value = nRealizacion
        oComm.Parameters.Add("@nGravament", SqlDbType.Money).Value = nGravament
        oComm.Parameters.Add("@nSaldo", SqlDbType.Money).Value = nSaldo
        oComm.Parameters.Add("@bGarantReal", SqlDbType.Bit).Value = bGarantReal
        oComm.Parameters.Add("@nTipoRealizac", SqlDbType.Int).Value = nTipoRealizacion
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 100).Value = sComentario
        oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 100).Value = sDireccion
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Garantia_Modificar(ByVal nCodGarant As Integer, _
                            ByVal dFecReg As Date, _
                            ByVal nTipoGarant As Integer, _
                            ByVal nClaseGarant As Integer, _
                            ByVal nTipoDoc As Integer, _
                            ByVal sNroDoc As String, _
                            ByVal nCodEmis As Integer, _
                            ByVal nMoneda As Integer, _
                            ByVal sDescripcion As String, _
                            ByVal sCodZona As String, _
                            ByVal nEstado As Integer, _
                            ByVal nTasacion As Double, _
                            ByVal nRealizacion As Double, _
                            ByVal nGravament As Double, _
                            ByVal nSaldo As Double, _
                            ByVal bGarantReal As Integer, _
                            ByVal nTipoRealizacion As Integer, _
                            ByVal sComentario As String, _
                            ByVal sDireccion As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Garantia_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@nTipoGarant", SqlDbType.Int).Value = nTipoGarant
        oComm.Parameters.Add("@nClaseGarant", SqlDbType.Int).Value = nClaseGarant
        oComm.Parameters.Add("@nTipoDoc", SqlDbType.Int).Value = nTipoDoc
        oComm.Parameters.Add("@cNroDoc", SqlDbType.VarChar, 50).Value = sNroDoc
        oComm.Parameters.Add("@nCodEmis", SqlDbType.Int).Value = nCodEmis
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Value = sDireccion
        oComm.Parameters.Add("@cCodZona", SqlDbType.Char, 12).Value = sCodZona
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = nEstado
        oComm.Parameters.Add("@nTasacion", SqlDbType.Money).Value = nTasacion
        oComm.Parameters.Add("@nRealizacion", SqlDbType.Money).Value = nRealizacion
        oComm.Parameters.Add("@nGravament", SqlDbType.Money).Value = nGravament
        oComm.Parameters.Add("@nSaldo", SqlDbType.Money).Value = nSaldo
        oComm.Parameters.Add("@bGarantReal", SqlDbType.Bit).Value = bGarantReal
        oComm.Parameters.Add("@nTipoRealizac", SqlDbType.Int).Value = nTipoRealizacion
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 100).Value = sComentario
        oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 100).Value = sDireccion
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function Garantia_Eliminar(ByVal nCodGarant As Integer) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Garantia_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        Garantia_Eliminar = oComm.Parameters("@pnResul").Value

        objCon = Nothing
        oComm = Nothing
    End Function

    Public Sub BuscaGarantia_x_Cliente(ByRef oCred As ComTipos.clsDataset, _
                ByVal sNomCliente As String, ByVal pnBusq As Integer)

        Dim oData As New ComTipos.clsDataset
        If pnBusq = 1 Then
            oData.CargaTabla(, , "Cred_BuscaGarantia_x_Cliente '" & sNomCliente & "'")
        End If
        If pnBusq = 2 Then
            oData.CargaTabla(, , "Cred_BuscaGarantia_x_DNI '" & sNomCliente & "'")
        End If
        If pnBusq = 3 Then
            oData.CargaTabla(, , "Cred_BuscaGarantia_x_Codigo '" & sNomCliente & "'")
        End If

        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Listar_TiposGarant_x_Clase(ByRef oCred As ComTipos.clsDataset, _
                ByVal nClaseGarant As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Listar_TiposGarant_x_Clase " & nClaseGarant)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Listar_Documentos_x_TipoGarant(ByRef oCred As ComTipos.clsDataset, _
            ByVal nTipoGarant As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Listar_Documentos_x_TipoGarant " & nTipoGarant)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub CargarInterfazGarant(ByRef oClase As ComTipos.clsDataset, _
                                ByRef oMoneda As ComTipos.clsDataset, _
                                ByRef oTDocum As ComTipos.clsDataset, _
                                ByRef oRelac As ComTipos.clsDataset, _
                                ByRef oEstado As ComTipos.clsDataset)

        Dim C As ComConstante.clsConstante
        C = New ComConstante.clsConstante
        oClase = C.RecuperaConstantes(2150)
        oMoneda = C.RecuperaConstantes(2090)
        oTDocum = C.RecuperaConstantes(2152)
        oRelac = C.RecuperaConstantes(2153)
        oEstado = C.RecuperaConstantes(2154)
        C = Nothing
    End Sub

    Public Sub RecuperaGarantPersonas(ByRef oCred As ComTipos.clsDataset, _
            ByVal nCodGarant As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaGarantPersonas " & nCodGarant)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RecuperaGarantDJ(ByRef oCred As ComTipos.clsDataset, _
            ByVal nCodGarant As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Recupera_GarantDJ " & nCodGarant)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub GarantDJ_Eliminar(ByVal nCodGarant As Integer, _
                                ByVal nItem As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GarantDJ_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = nItem
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GarantDJ_Insertar(ByVal nCodGarant As Integer, _
                        ByVal nItem As Integer, _
                        ByVal sDescripcion As String, _
                        ByVal nCantidad As Integer, _
                        ByVal nPrecioUnit As Double, _
                        ByVal nDocTpo As Integer, _
                        ByVal sNroDoc As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GarantDJ_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = nItem
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Value = sDescripcion
        oComm.Parameters.Add("@nCantidad", SqlDbType.Int).Value = nCantidad
        oComm.Parameters.Add("@nPrecioUnit", SqlDbType.Money).Value = nPrecioUnit
        oComm.Parameters.Add("@nDocTpo", SqlDbType.Int).Value = nDocTpo
        oComm.Parameters.Add("@cNroDoc", SqlDbType.VarChar, 20).Value = sNroDoc
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Recupera_Garantia(ByRef oCred As ComTipos.clsDataset, _
        ByVal nCodGarant As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Recupera_Garantia " & nCodGarant)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Garantia_Grabar(ByVal nCodGarant As Integer, _
                            ByVal dFecReg As Date, _
                            ByVal nTipoGarant As Integer, _
                            ByVal nClaseGarant As Integer, _
                            ByVal nTipoDoc As Integer, _
                            ByVal sNroDoc As String, _
                            ByVal nCodEmis As Integer, _
                            ByVal nMoneda As Integer, _
                            ByVal sDescripcion As String, _
                            ByVal sCodZona As String, _
                            ByVal nEstado As Integer, _
                            ByVal nTasacion As Double, _
                            ByVal nRealizacion As Double, _
                            ByVal nDisponible As Double, _
                            ByVal nSaldo As Double, _
                            ByVal bGarantReal As Integer, _
                            ByVal nTipoRealizacion As Integer, _
                            ByVal sComentario As String, _
                            ByVal sDireccion As String, _
                            ByVal oGarantPers As ComTipos.clsDataset, _
                            ByVal oGarantDJ As ComTipos.clsDataset, _
                            ByVal bNuevo As Integer, ByVal dFecEval As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta

        'ARCV 11-05-2011
        'oComm = New SqlCommand("Cred_Garantia_Grabar", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_Garantia_Grabar")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@nTipoGarant", SqlDbType.Int).Value = nTipoGarant
        oComm.Parameters.Add("@nClaseGarant", SqlDbType.Int).Value = nClaseGarant
        oComm.Parameters.Add("@nTipoDoc", SqlDbType.Int).Value = nTipoDoc
        oComm.Parameters.Add("@cNroDoc", SqlDbType.VarChar, 50).Value = sNroDoc
        oComm.Parameters.Add("@nCodEmis", SqlDbType.Int).Value = nCodEmis
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Value = sDescripcion
        oComm.Parameters.Add("@cCodZona", SqlDbType.Char, 12).Value = sCodZona
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = nEstado
        oComm.Parameters.Add("@nTasacion", SqlDbType.Money).Value = nTasacion
        oComm.Parameters.Add("@nRealizacion", SqlDbType.Money).Value = nRealizacion
        oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = nDisponible
        oComm.Parameters.Add("@nSaldo", SqlDbType.Money).Value = nSaldo
        oComm.Parameters.Add("@bGarantReal", SqlDbType.Bit).Value = bGarantReal
        oComm.Parameters.Add("@nTipoRealizac", SqlDbType.Int).Value = nTipoRealizacion
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 100).Value = sComentario
        oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 100).Value = sDireccion
        oComm.Parameters.Add("@oGarantPers", SqlDbType.Xml).Value = oGarantPers.DevuelveXML
        oComm.Parameters.Add("@oGarantDJ", SqlDbType.Xml).Value = oGarantDJ.DevuelveXML
        oComm.Parameters.Add("@bNuevo", SqlDbType.Bit).Value = bNuevo
        oComm.Parameters.Add("@dFecEval", SqlDbType.DateTime).Value = dFecEval

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CargarInterfazGarantReal(ByRef oTipoV As ComTipos.clsDataset, _
                            ByRef oEstadoH As ComTipos.clsDataset)

        Dim C As ComConstante.clsConstante
        C = New ComConstante.clsConstante
        oTipoV = C.RecuperaConstantes(2155)
        oEstadoH = C.RecuperaConstantes(2156)
        C = Nothing
    End Sub

    Public Sub Recupera_GarantiaReal(ByRef oCred As ComTipos.clsDataset, _
    ByVal nCodGarant As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Recupera_GarantReal " & nCodGarant)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub GarantReal_Insertar(ByVal nCodGarant As Integer, _
                        ByVal dFecha As Date, _
                        ByVal dFecReg As Date, _
                        ByVal nCodTasador As Integer, _
                        ByVal sNroPartida As String, _
                        ByVal sNroFoja As String, _
                        ByVal sNroTomo As String, _
                        ByVal nTipoVivienda As Integer, _
                        ByVal nEstadoHip As Integer, _
                        ByVal sPlacaVehi As String, _
                        ByVal sModeloVehi As String, _
                        ByVal sMarcaVehi As String, _
                        ByVal bGravamen As Integer, _
                        ByVal sPersGravamen As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GarantReal_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@dFecRegistro", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@nCodTasador", SqlDbType.Int).Value = nCodTasador
        oComm.Parameters.Add("@cNroPartida", SqlDbType.VarChar, 20).Value = sNroPartida
        oComm.Parameters.Add("@cNroFoja", SqlDbType.VarChar, 20).Value = sNroFoja
        oComm.Parameters.Add("@cNroTomo", SqlDbType.VarChar, 20).Value = sNroTomo
        oComm.Parameters.Add("@nTipoVivienda", SqlDbType.Int).Value = nTipoVivienda
        oComm.Parameters.Add("@nEstadoHipo", SqlDbType.Int).Value = nEstadoHip
        oComm.Parameters.Add("@cPlacaVehi", SqlDbType.VarChar, 20).Value = sPlacaVehi
        oComm.Parameters.Add("@cModeloVehi", SqlDbType.VarChar, 20).Value = sModeloVehi
        oComm.Parameters.Add("@cMarcaVehi", SqlDbType.VarChar, 20).Value = sMarcaVehi

        oComm.Parameters.Add("@bGravamen", SqlDbType.Bit).Value = bGravamen
        oComm.Parameters.Add("@cPersGravamen", SqlDbType.VarChar, 20).Value = sPersGravamen

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GarantReal_Modificar(ByVal nCodGarant As Integer, _
                    ByVal dFecha As Date, _
                    ByVal dFecReg As Date, _
                    ByVal nCodTasador As Integer, _
                    ByVal sNroPartida As String, _
                    ByVal sNroFoja As String, _
                    ByVal sNroTomo As String, _
                    ByVal nTipoVivienda As Integer, _
                    ByVal nEstadoHip As Integer, _
                    ByVal sPlacaVehi As String, _
                    ByVal sModeloVehi As String, _
                    ByVal sMarcaVehi As String, _
                        ByVal bGravamen As Integer, _
                        ByVal sPersGravamen As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GarantReal_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@dFecRegistro", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@nCodTasador", SqlDbType.Int).Value = nCodTasador
        oComm.Parameters.Add("@cNroPartida", SqlDbType.VarChar, 20).Value = sNroPartida
        oComm.Parameters.Add("@cNroFoja", SqlDbType.VarChar, 20).Value = sNroFoja
        oComm.Parameters.Add("@cNroTomo", SqlDbType.VarChar, 20).Value = sNroTomo
        oComm.Parameters.Add("@nTipoVivienda", SqlDbType.Int).Value = nTipoVivienda
        oComm.Parameters.Add("@nEstadoHipo", SqlDbType.Int).Value = nEstadoHip
        oComm.Parameters.Add("@cPlacaVehi", SqlDbType.VarChar, 20).Value = sPlacaVehi
        oComm.Parameters.Add("@cModeloVehi", SqlDbType.VarChar, 20).Value = sModeloVehi
        oComm.Parameters.Add("@cMarcaVehi", SqlDbType.VarChar, 20).Value = sMarcaVehi

        oComm.Parameters.Add("@bGravamen", SqlDbType.Bit).Value = bGravamen
        oComm.Parameters.Add("@cPersGravamen", SqlDbType.VarChar, 20).Value = sPersGravamen

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GarantReal_Eliminar(ByVal nCodGarant As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GarantReal_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGarant", SqlDbType.Int).Value = nCodGarant

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Buscar_Personas_x_Credito(ByRef oCred As ComTipos.clsDataset, _
                ByVal nCodCred As Integer, _
                ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Buscar_Personas_x_Credito " & nCodCred & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Buscar_Garantias_x_Persona(ByRef oCred As ComTipos.clsDataset, _
            ByVal nCodPers As Integer, _
            ByVal nMoneda As Integer, _
            ByVal nCodCred As Integer, _
                ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Buscar_Garantias_x_Persona " & nCodPers & "," & nMoneda & "," & nCodCred & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Gravar_Garantia(ByVal nCodCred As Integer, _
                                ByVal nCodAge As Integer, _
                                ByVal dFecha As Date, _
                                ByVal oGravament As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Cred_Gravar_Garantia", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_Gravar_Garantia")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@oGravament", SqlDbType.Xml).Value = oGravament.DevuelveXML

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CastigarCredito(ByVal nCodCred As Integer, _
                                ByVal nCodAge As Integer, _
                                ByVal nCodAgeTrans As Integer, _
                                ByVal dFecha As Date, _
                                ByVal sUser As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CastigarCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodAgeTrans", SqlDbType.Int).Value = nCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@cUser", SqlDbType.Char, 4).Value = sUser
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub BuscarCreditos_x_DiasAtraso(ByRef oCred As ComTipos.clsDataset, _
                ByVal nDiasAtraso As Integer, _
                Optional ByVal nEstado As Integer = 0)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarCreditos_x_DiasAtraso " & nDiasAtraso & "," & nEstado)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub TransJudicial(ByVal dFecha As Date, _
                            ByVal bDemanda As Integer, _
                            ByVal oCreditos As ComTipos.clsDataset, _
                            ByVal sUser As String, _
                            ByVal nCodAgeTrans As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_TransJudicial", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@bDemanda", SqlDbType.Bit).Value = bDemanda
        oComm.Parameters.Add("@oCreditos", SqlDbType.Xml).Value = oCreditos.DevuelveXML
        oComm.Parameters.Add("@cUser", SqlDbType.Char, 4).Value = sUser
        oComm.Parameters.Add("@nCodAgeTrans", SqlDbType.Int).Value = nCodAgeTrans

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GarantPorcen_Buscar(ByRef oCred As ComTipos.clsDataset, _
        ByVal nTipoGarant As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_GarantPorcen_Buscar " & nTipoGarant)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ValidaCuentasAsiento(ByRef oData As ComTipos.clsDataset, _
                                                                ByVal dFecha As String, _
                                                                ByVal nCodAge As Integer)

        Dim oDatos As New ComTipos.clsDataset
        oDatos.CargaTabla(, , "Cred_ValidaCuentasAsiento '" & dFecha & "'," & nCodAge)
        oData = oDatos
        oDatos = Nothing
    End Sub

    Public Sub BuscarExtornosProceso(ByRef oData As ComTipos.clsDataset, _
                                                                ByVal dFecha As String, _
                                                                ByVal nTipoProceso As Integer)

        Dim oDatos As New ComTipos.clsDataset
        oDatos.CargaTabla(, , "Cred_BuscarExtornosProceso '" & dFecha & "'," & nTipoProceso)
        oData = oDatos
        oDatos = Nothing
    End Sub

    Public Sub Extorno_TransJudicial(ByVal nNroTrans As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Extorno_TransJudicial", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTrans

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ExtornoCastigoCredito(ByVal nNroTrans As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ExtornoCastigoCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTrans

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarCrediPago_WIN(ByRef oCreditos As ComTipos.clsDataset, _
                           ByVal pnCodBanco As Integer, _
                           ByRef psError As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'objCon = New ComConecta.clsConecta

        With oCreditos
            .Inicio()

            While Not .EOF
                objCon = New ComConecta.clsConecta

                oComm = New SqlCommand("Cred_GrabaCrediPago", objCon.DameConexion)
                'oComm.CommandTimeout = 0

                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = .ObtenerDato("CodAge")
                oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = .ObtenerDato("CodCred")
                oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = CDate(.ObtenerDato("Fecha"))
                oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = CDbl(.ObtenerDato("Monto"))
                oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = pnCodBanco

                oComm.Parameters.Add("@cERROR", SqlDbType.VarChar, 100)
                oComm.Parameters.Item("@cERROR").Direction = ParameterDirection.Output

                oComm.ExecuteNonQuery()

                psError = oComm.Parameters.Item("@cERROR").Value

                oComm.Dispose()
                oComm = Nothing

                'ARCV 05-11-2008
                'If psError = "MAYORDEUDA" Then
                '    'oComm.Dispose()
                '    'oComm = Nothing
                '    objCon.CierraConexion()
                '    objCon = Nothing
                '    Exit Sub
                'End If


                objCon.CierraConexion()
                objCon = Nothing

                .Siguiente()
            End While
        End With

        'objCon.CierraConexion()
        'objCon = Nothing

    End Sub

    'Public Sub GrabarCrediPago_WIN(ByRef oCreditos As ComTipos.clsDataset, _
    '                      ByVal pnCodBanco As Integer, _
    '                      ByRef psError As String)

    '    Dim objCon As ComConecta.clsConecta
    '    Dim oComm As SqlCommand

    '    objCon = New ComConecta.clsConecta

    '    With oCreditos
    '        .Inicio()


    '        While Not .EOF
    '            'objCon = New ComConecta.clsConecta

    '            oComm = New SqlCommand("Cred_GrabaCrediPago", objCon.DameConexion)
    '            oComm.CommandType = CommandType.StoredProcedure
    '            oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = .ObtenerDato("CodAge")
    '            oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = .ObtenerDato("CodCred")
    '            oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = CDate(.ObtenerDato("Fecha"))
    '            oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = CDbl(.ObtenerDato("Monto"))
    '            oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = pnCodBanco

    '            oComm.Parameters.Add("@cERROR", SqlDbType.VarChar, 100)
    '            oComm.Parameters.Item("@cERROR").Direction = ParameterDirection.Output

    '            oComm.ExecuteNonQuery()

    '            psError = oComm.Parameters.Item("@cERROR").Value
    '            oComm.Dispose()
    '            oComm = Nothing

    '            If psError = "MAYORDEUDA" Then
    '                objCon = Nothing
    '                Exit Sub
    '            End If

    '            'objCon = Nothing

    '            .Siguiente()
    '        End While
    '    End With
    '    objCon = Nothing

    'End Sub

    Public Sub ReporteCarteraRefinanciada(ByRef oCred As ComTipos.clsDataset, _
                    ByVal pnTipoCambio As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ReporteCarteraRefinanciada " & pnTipoCambio)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RegistraComentario(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal psComen As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RegistraComentario", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@sComentario", SqlDbType.Text).Value = psComen

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing
    End Sub

    Public Sub RecuperaComentario(ByRef oCred As ComTipos.clsDataset, _
                ByVal pnCodCred As Integer, _
                ByVal pnCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaComentario " & pnCodCred & "," & pnCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub ReporteCredAtrasoM90Dias(ByRef oCred As ComTipos.clsDataset, _
                ByVal pnTipoCambio As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ReporteCredAtrasoM90Dias " & pnTipoCambio)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Anexo02_A(ByRef oCred As ComTipos.clsDataset, _
                ByVal pnTipoCambio As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Anexo02_A " & pnTipoCambio)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Anexo02_C(ByRef oCred As ComTipos.clsDataset, _
            ByVal pnTipoCambio As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Anexo02_C " & pnTipoCambio)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Anexo5G_Datos(ByRef oCred As ComTipos.clsDataset, _
                    ByVal pnTipoCambio As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Anexo5G_Datos " & pnTipoCambio)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RecuperaCreditoDatos(ByRef oCred As ComTipos.clsDataset, _
                        ByVal nCodCred As Integer, _
                        ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaCreditoDatos " & nCodCred & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RecuperaCredPersonas(ByRef oCred As ComTipos.clsDataset, _
                        ByVal nCodCred As Integer, _
                        ByVal nCodAge As Integer, _
                        ByVal nRelacion As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaCredPersonas " & nCodCred & "," & nCodAge & "," & nRelacion)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub BuscarGarantiaGrupal(ByRef oCred As ComTipos.clsDataset, _
                    ByVal nGrupo As Integer, _
                    ByVal nCiclo As Integer, _
                    ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarGarantiaGrupal " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Graba_GarantiaGrupal(ByVal dFecha As Date, _
                ByVal nMoneda As Integer, _
                ByVal oDatos As ComTipos.clsDataset, _
                ByVal nCodGrupo As Integer, _
                ByVal nCiclo As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Cred_Graba_GarantiaGrupal", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_Graba_GarantiaGrupal")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@oDatos", SqlDbType.Xml).Value = oDatos.DevuelveXML
        oComm.Parameters.Add("@nCodGrupo", SqlDbType.Int).Value = nCodGrupo
        oComm.Parameters.Add("@nCiclo", SqlDbType.Int).Value = nCiclo

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing
    End Sub

    Public Function RecuperaZona(ByVal nCodPers As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Pers_RecuperaZona " & nCodPers)
        Return oData

    End Function

    Public Sub BuscarCronogramaGrupal(ByRef oCred As ComTipos.clsDataset, _
                        ByVal nGrupo As Integer, _
                        ByVal nCiclo As Integer, _
                        ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarCronogramaGrupal " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Function BuscaTitularGarantia(ByVal nCodGarant As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscaTitularGarantia " & nCodGarant)
        Return oData
        oData = Nothing
    End Function

    Public Function BuscarExtornos_Grupo(ByVal nCodGrupo As Integer, _
                                        ByVal nCiclo As Integer, _
                                        ByVal nCodAge As Integer, _
                                        ByVal dFecha As Date, _
                                        ByVal sUser As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarExtornos_Grupo " & nCodGrupo & "," & nCiclo & "," & nCodAge & ",'" & dFecha & "','" & sUser & "'")
        Return oData
        oData = Nothing
    End Function

    Public Function BuscarExtornosDesemb_Grupo(ByVal nCodGrupo As Integer, _
                                    ByVal nCiclo As Integer, _
                                    ByVal dFecha As Date, _
                                    ByVal sUser As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarExtornosDesemb_Grupo " & nCodGrupo & "," & nCiclo & ",'" & dFecha & "','" & sUser & "'")
        Return oData
        oData = Nothing
    End Function

    Public Sub GrabarCrediPagoGrupo_WIN(ByRef oCreditos As ComTipos.clsDataset, _
                                ByVal pnCodBanco As Integer, _
                                ByRef psError As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta

        With oCreditos
            .Inicio()
            While Not .EOF
                oComm = New SqlCommand("Cred_GrabaCrediPagoGrupo", objCon.DameConexion)
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar, 3).Value = .ObtenerDato("CodAge")
                oComm.Parameters.Add("@cCod_Grupo", SqlDbType.VarChar, 10).Value = .ObtenerDato("CodGrupo")
                oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = CDate(.ObtenerDato("Fecha"))
                oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = CDbl(.ObtenerDato("Monto"))
                oComm.Parameters.Add("@nTipoGrupo", SqlDbType.Int).Value = CInt(.ObtenerDato("TipoGrupo"))
                oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = pnCodBanco

                oComm.Parameters.Add("@cCiclo", SqlDbType.VarChar, 3).Value = .ObtenerDato("Ciclo")

                oComm.Parameters.Add("@cCod_AgeAnt", SqlDbType.VarChar, 3).Value = .ObtenerDato("CodAgeAnt") 'ARCV 03-01-2009

                oComm.Parameters.Add("@cERROR", SqlDbType.VarChar, 100)
                oComm.Parameters.Item("@cERROR").Direction = ParameterDirection.Output

                oComm.ExecuteNonQuery()

                psError = oComm.Parameters.Item("@cERROR").Value

                oComm = Nothing

                'If psError <> "" Then
                'objCon = Nothing
                'Exit Sub
                'End If

                .Siguiente()
            End While
        End With
        objCon = Nothing

    End Sub

    Public Sub Desgravamen(ByVal nCodCred As Integer, _
            ByVal nCodAge As Integer, _
            ByVal dFecha As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Desgravamen", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing
    End Sub

    Public Function CargaDatosSolicitudGrupal_Duplic(ByVal pnCodGrp As Integer, _
                                        ByVal pnCiclo As Integer) As ComTipos.ClsWCredSolGrp
        Dim D As ComTipos.ClsWCredSolGrp
        D = New ComTipos.ClsWCredSolGrp

        'D = NuevoCreditoGrupal(pnCodAge, pdFecSis)
        D.Integrantes.CargaTabla(, , "Cred_RecuperaSolGrpIntegrantes_Dupli " & pnCodGrp & "," & pnCiclo)
        If D.Integrantes.nNumFilas > 0 Then
            D.Integrantes.Inicio()
            Do While Not D.Integrantes.EOF
                D.Integrantes.AsignarDato(D.SeparaActividad(D.Integrantes.ObtenerDato("cActividad")), "cActividad")
                D.Integrantes.Siguiente()
            Loop
            D.Integrantes.Seek(1)
            D.Datos.CargaTabla(, , "CredW_RecuperaDatosSolGrp " & D.Integrantes.ObtenerDato("nCodCred") & "," & D.Integrantes.ObtenerDato("nCodAge"))
        End If
        CargaDatosSolicitudGrupal_Duplic = D
    End Function

    Public Function BuscarCodigoGrupo(ByVal pnCodAge As Integer, _
                                        ByVal pnCodGrupo As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_BuscarCodigoGrupo " & pnCodGrupo & "," & pnCodAge)
        BuscarCodigoGrupo = C
        C = Nothing
    End Function

    Public Sub GrabarPrepago(ByVal sCodExpe As String, _
                              ByVal sCodAge As String, _
                              ByVal sCodAgeTrans As String, _
                              ByVal dFecha As Date, _
                              ByVal nMonto As Double, _
                              ByVal nCodOpe As Integer, _
                              ByVal nCodBanco As Integer, _
                              ByVal nITF As Double, _
                          ByRef nOCAPITAL As Double, _
                          ByRef nOINTERES As Double, _
                          ByRef nOGASTOS As Double, _
                          ByRef nOMORA As Double, _
                          ByRef nOIGV As Double, _
                          ByRef nOSALDO As Double, _
                          ByRef nOPROXCUOTA As Integer, _
                          ByRef nONROTRANS As Integer, _
                          ByVal sUser As String, _
                          ByRef sOAGEDIR As String, _
                          ByRef sOAGETEL As String, _
                          ByRef sODESCOPE As String, _
                          ByRef sPCOMENTB As String, _
                          ByRef sSCOMENTB As String, _
                          Optional ByVal sNumCheque As String = "", _
                          Optional ByRef sProxVenc As String = "", _
                          Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        'ARCV 11-05-2011
        'oComm = New SqlClient.SqlCommand("Cred_PagoCuotas_Prepago", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_PagoCuotas_Prepago")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output
        '---------------

        'oComm.ExecuteNonQuery()
        'ARCV 11-05-2011
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function Recupera_CredCtaAnt(ByVal nCodCred As Integer, _
                                ByVal nCodAge As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Recupera_CredCtaAnt " & nCodCred & "," & nCodAge)
        Return oData
        oData = Nothing
    End Function

    Public Function Recupera_CodGrupoAgeAnt(ByVal sCodGrupo As String, _
                                ByVal nCiclo As Integer, _
                                ByVal nCodAge As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Recupera_CodGrupoAgeAnt '" & sCodGrupo & "'," & nCiclo & "," & nCodAge)
        Return oData
        oData = Nothing
    End Function

    Public Function RecuperaAsientoContNEGOCIO(ByVal pdFecha As Date, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim oDato As ComTipos.clsDataset
        oDato = New ComTipos.clsDataset
        'oDato.CargaTabla(, , "Cred_RecuperaAsientoContNEGOCIO '" & Format(pdFecha, "dd/MM/yyyy") & "'," & pnCodAge, 3)
        oDato.CargaTabla(, , "Cred_RecuperaAsientoContNEGOCIO '" & Format(pdFecha, "dd/MM/yyyy") & "'," & pnCodAge)
        RecuperaAsientoContNEGOCIO = oDato
        oDato = Nothing
    End Function

    Public Sub EliminaAsientoMigrados(ByVal dFechaI As Date, _
                                    ByVal dFechaF As Date, _
                                ByVal nCodAge As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Finan_EliminaAsientoMigrados", objCon.DameConexionFinanc)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFechaI", SqlDbType.DateTime).Value = dFechaI
        oComm.Parameters.Add("@dFechaF", SqlDbType.DateTime).Value = dFechaF
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing
    End Sub

    Public Sub MigraAsientoContable(ByVal dFecha As Date, _
                        ByVal nCodAge As Integer, _
                        ByVal oAsiento As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Finan_MigraAsientoContable", objCon.DameConexionFinanc)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nMovMonto", SqlDbType.Money).Value = oAsiento.DevolverSumaColumna("nDebe")
        oComm.Parameters.Add("@oAsiento", SqlDbType.Xml).Value = oAsiento.DevuelveXML

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing
    End Sub

    Public Sub GrabarCondonacionCredito(ByVal nCodCred As Integer, _
                    ByVal nCodAge As Integer, _
                    ByVal dFecha As Date, _
                    ByVal nDiasAtraso As Integer, _
                    ByVal nCapital As Double, _
                    ByVal nInteres As Double, _
                    ByVal nMora As Double, _
                    ByVal nGasto As Double, _
                    ByVal sUser As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CondonarCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nDiasAtraso", SqlDbType.Int).Value = nDiasAtraso
        oComm.Parameters.Add("@nCapital", SqlDbType.Money).Value = nCapital
        oComm.Parameters.Add("@nInteres", SqlDbType.Money).Value = nInteres
        oComm.Parameters.Add("@nMora", SqlDbType.Money).Value = nMora
        oComm.Parameters.Add("@nGastos", SqlDbType.Money).Value = nGasto
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ExtornoAprobacion(ByVal nCodCred As Integer, _
                ByVal nCodAge As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ExtornoAprobacion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function ValidaIntegracionDATA(ByVal pnTipoValida As Integer) As ComTipos.clsDataset
        Dim oDato As ComTipos.clsDataset
        oDato = New ComTipos.clsDataset
        oDato.CargaTabla(, , "Cred_ValidaIntegracionDATA " & pnTipoValida, , , 0)
        ValidaIntegracionDATA = oDato
        oDato = Nothing
    End Function

    Public Sub CorrigeIntegracionDATA(ByVal pnTipoValida As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CorrigeIntegracionDATA", objCon.DameConexion)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nTipoValida", SqlDbType.Int).Value = pnTipoValida

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function RecuperaConsultaCredDatosFIDEICOMISO(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsCreConsulta
        Dim D As ComTipos.ClsCreConsulta
        D = New ComTipos.ClsCreConsulta
        D.Data.CargaTabla(, , "DBCORE..Cred_RecuperaDatosConsCred " & pnCodCred & "," & pnCodAge)
        D.CuotaPag.CargaTabla(, , "DBCORE..Cred_RecuperaConsCred_CuotasPag " & pnCodCred & "," & pnCodAge)
        D.CuotaPend.CargaTabla(, , "DBCORE..Cred_RecuperaConsCred_CuotasPend " & pnCodCred & "," & pnCodAge)
        D.DataDeudCal.CargaTabla(, , "DBCORE..Cred_RecuperaConsCred_DeudaTotalCal " & pnCodCred & "," & pnCodAge)
        D.Kardex.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredKardex " & pnCodCred & "," & pnCodAge)
        D.DataDeudaFecha.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredDeudaFecha " & pnCodCred & "," & pnCodAge)
        D.Garantias.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredGarantia " & pnCodCred & "," & pnCodAge)
        D.DataRelaPers.CargaTabla(, , "DBCORE..Cred_RecuperaConsCredRelaPers " & pnCodCred & "," & pnCodAge)
        D.DatosCalendario.CargaTabla(, , "DBCORE..Cred_RecuperaDatosBasicosCalendario " & pnCodAge & "," & pnCodCred)
        D.CalendDesembolsos.CargaTabla(, , "DBCORE..Cred_RecuperaCalendDesembolsos " & pnCodAge & "," & pnCodCred)
        RecuperaConsultaCredDatosFIDEICOMISO = D
        D = Nothing
    End Function

    Public Sub BuscarDatosCreditoFIDEICOMISO(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByRef oTipoCred As ComTipos.clsCredPagoCuotas, _
                                    ByVal dFecha As Date, _
                                    ByVal nEstado As Integer)

        Dim oData As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante

        oData.CargaTabla(, , "DBCORE..CredW_BuscaExpediente '" & sCodAge & "','" & sCodExpe & "','" & dFecha & "'," & nEstado)
        'Return oData
        oTipoCred.Datos = oData
        oTipoCred.Bancos = oCons.RecuperaConstantes(2000)
        oData = Nothing
        oCons = Nothing
    End Sub

    Public Sub BuscarDatosCreditoEDPYME(ByVal sCodExpe As String, _
                                ByVal sCodAge As String, _
                                ByRef oTipoCred As ComTipos.clsCredPagoCuotas, _
                                ByVal dFecha As Date, _
                                ByVal nEstado As Integer)

        Dim oData As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante

        oData.CargaTabla(, , "DBSOLIDA..CredW_BuscaExpediente '" & sCodAge & "','" & sCodExpe & "','" & dFecha & "'," & nEstado)
        'Return oData
        oTipoCred.Datos = oData
        oTipoCred.Bancos = oCons.RecuperaConstantes(2000)
        oData = Nothing
        oCons = Nothing
    End Sub

    Public Sub BuscarDetalleCuotasFIDEICOMISO(ByVal sCodExpe As String, _
                                ByVal sCodAge As String, _
                                ByRef oPag As ComTipos.clsCredPagoCuotas)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBCORE..Cred_DetalleCuotas '" & sCodAge & "','" & sCodExpe & "'")
        oPag.Cuotas = oData
        oData = Nothing
    End Sub

    Public Sub BuscarDetalleCuotasEDPYME(ByVal sCodExpe As String, _
                            ByVal sCodAge As String, _
                            ByRef oPag As ComTipos.clsCredPagoCuotas)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBSOLIDA..Cred_DetalleCuotas '" & sCodAge & "','" & sCodExpe & "'")
        oPag.Cuotas = oData
        oData = Nothing
    End Sub

    Public Sub GrabarPagoCuotas_WINFIDEICOMISO(ByVal sCodExpe As String, _
                          ByVal sCodAge As String, _
                          ByVal sCodAgeTrans As String, _
                          ByVal dFecha As Date, _
                          ByVal nMonto As Double, _
                          ByVal nCodOpe As Integer, _
                          ByVal nCodBanco As Integer, _
                          ByVal nITF As Double, _
                      ByRef nOCAPITAL As Double, _
                      ByRef nOINTERES As Double, _
                      ByRef nOGASTOS As Double, _
                      ByRef nOMORA As Double, _
                      ByRef nOIGV As Double, _
                      ByRef nOSALDO As Double, _
                      ByRef nOPROXCUOTA As Integer, _
                      ByRef nONROTRANS As Integer, _
                      ByVal sUser As String, _
                      ByRef sOAGEDIR As String, _
                      ByRef sOAGETEL As String, _
                      ByRef sODESCOPE As String, _
                      ByRef sPCOMENTB As String, _
                      ByRef sSCOMENTB As String, _
                      Optional ByVal sNumCheque As String = "", _
                      Optional ByRef sProxVenc As String = "", _
                      Optional ByVal nPagoCta As Integer = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("DBCORE..Cred_PagoCuotas_FIDEICOMISO", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarPagoCuotas_FIDEICOMISO(ByVal sCodExpe As String, _
                          ByVal sCodAge As String, _
                          ByVal sCodAgeTrans As String, _
                          ByVal dFecha As Date, _
                          ByVal nMonto As Double, _
                          ByVal nCodOpe As Integer, _
                          ByVal nCodBanco As Integer, _
                          ByVal nITF As Double, _
                      ByRef nOCAPITAL As Double, _
                      ByRef nOINTERES As Double, _
                      ByRef nOGASTOS As Double, _
                      ByRef nOMORA As Double, _
                      ByRef nOIGV As Double, _
                      ByRef nOSALDO As Double, _
                      ByRef nOPROXCUOTA As Integer, _
                      ByRef nONROTRANS As Integer, _
                      ByVal sUser As String, _
                      ByRef sOAGEDIR As String, _
                      ByRef sOAGETEL As String, _
                      ByRef sODESCOPE As String, _
                      ByRef sPCOMENTB As String, _
                      ByRef sSCOMENTB As String, _
                      Optional ByVal sNumCheque As String = "", _
                      Optional ByRef sProxVenc As String = "", _
                      Optional ByVal nPagoCta As Integer = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("DBCORE..Cred_PagoCuotas", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub GrabarPagoCuotas_WINEDPYME(ByVal sCodExpe As String, _
                      ByVal sCodAge As String, _
                      ByVal sCodAgeTrans As String, _
                      ByVal dFecha As Date, _
                      ByVal nMonto As Double, _
                      ByVal nCodOpe As Integer, _
                      ByVal nCodBanco As Integer, _
                      ByVal nITF As Double, _
                  ByRef nOCAPITAL As Double, _
                  ByRef nOINTERES As Double, _
                  ByRef nOGASTOS As Double, _
                  ByRef nOMORA As Double, _
                  ByRef nOIGV As Double, _
                  ByRef nOSALDO As Double, _
                  ByRef nOPROXCUOTA As Integer, _
                  ByRef nONROTRANS As Integer, _
                  ByVal sUser As String, _
                  ByRef sOAGEDIR As String, _
                  ByRef sOAGETEL As String, _
                  ByRef sODESCOPE As String, _
                  ByRef sPCOMENTB As String, _
                  ByRef sSCOMENTB As String, _
                  Optional ByVal sNumCheque As String = "", _
                  Optional ByRef sProxVenc As String = "", _
                  Optional ByVal nPagoCta As Integer = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("DBSOLIDA..Cred_PagoCuotas_EDPYME", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub BuscarDesembolsoEDPYME(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByRef oDesemb As ComTipos.clsCredDesembolso)
        Dim oData As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante

        oData.CargaTabla(, , "DBSOLIDA..CredW_BuscaDesembolso '" & sCodAge & "','" & sCodExpe & "'")
        oDesemb.Datos = oData
        oDesemb.Bancos = oCons.RecuperaConstantes(2000)
        oData = Nothing
    End Sub

    Public Function CreditoAutorizadoEDPYME(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("DBSOLIDA..Cred_CreditoAutorizado ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        CreditoAutorizadoEDPYME = oComm.Parameters("@pnResul").Value
        oComm = Nothing
        objCon = Nothing
    End Function

    Public Function ObtieneGastoTotalCalNewEDPYME(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Double

        Dim R As ComTipos.clsDataset
        R = New ComTipos.clsDataset
        R.CargaTabla(, , "DBSOLIDA..Cred_RecuperaGastosALDesemb " & pnCodAge & "," & pnCodCred)

        ObtieneGastoTotalCalNewEDPYME = CDbl(R.ObtenerDato("nValorGasto"))
        R = Nothing
    End Function

    Public Sub DesembolsarCredito_WINEDPYME(ByVal sCodExpe As String, _
                                ByVal sCodAge As String, _
                                ByVal sCodAgeTrans As String, _
                                ByVal dFecha As Date, _
                                ByVal nPrestamo As Double, _
                                ByVal nCodOpe As Integer, _
                                ByVal nCodBanco As Integer, _
                                ByVal nITF As Double, _
                                ByVal nCuotas As Integer, _
                                ByVal nTasa As Double, _
                                ByVal nDia As Integer, _
                                ByVal nDiasGracia As Integer, _
                                ByVal nTipoGracia As Integer, _
                                ByVal nGastos As Double, _
                            ByRef dOFecVenc As Date, _
                            ByRef cOPagare As String, _
                            ByRef cODocEmi As String, _
                            ByRef cODomEmi As String, _
                            ByRef cOConyuge As String, _
                            ByRef cODocCon As String, _
                            ByRef cODomCon As String, _
                        ByRef cODpto As String, _
                        ByRef cOProv As String, _
                        ByRef cODist As String, _
                        ByRef cOLinea As String, _
                        ByRef cOMoneda As String, _
                        ByRef bODNI As Boolean, _
                        ByVal sUser As String, _
                        ByRef cOCodBCP As String, _
                        ByRef cOCodBN As String, _
                        ByVal nProd As Integer, _
                        ByVal nSubProd As Integer, _
                        ByRef cOGrupo As String, _
                        ByVal nPlazo As Integer, _
                        ByVal nTipoPeriodo As Integer, _
                        ByRef sPCOMENTC As String, _
                          ByRef sSCOMENTC As String, _
                          Optional ByVal sNumCheque As String = "", _
                          Optional ByVal sRUC As String = "", _
                          Optional ByRef sZona As String = "", Optional ByVal nValorITF As Double = 0, Optional ByVal pnSoloIntereses As Integer = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim sGastos As String
        Dim sCalend As String
        Dim oTemp As Object
        'Dim MatGastosCol As Object
        Dim MatCalendCol As Object
        Dim dsCalend As New ComTipos.clsDataset
        Dim dsGastos As New ComTipos.clsDataset
        Dim dsCamp As New ComTipos.clsDataset
        Dim dsCredDesPar As New ComTipos.clsDataset
        Dim dsCalenDesPar As New ComTipos.clsDataset
        Dim dsCalDesembRenov As New ComTipos.clsDataset
        Dim oCalend As New ComTipos.ClsCalendario
        Dim i As Integer
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset
        Dim MatCalenCuotasDesemParDetalle As ComTipos.clsDataset
        Dim nPrestamoTempo As Double
        Dim nTasaDesgravamen As Double

        dsCamp.CargaTabla(, , "DBSOLIDA..CredW_RecuperaCreditoDatos " & sCodExpe & "," & sCodAge)

        dsCredDesPar.CargaTabla(, , "DBSOLIDA..Cred_RecuperaDatosDesemParcial " & sCodAge & "," & sCodExpe)
        dsCalenDesPar.CargaTabla(, , "DBSOLIDA..Cred_RecuperaCalDesParcial " & sCodAge & "," & sCodExpe)


        'ARCV 21-05-2009
        nMontoGasto = ObtieneGastoTotalCalNewEDPYME(CInt(sCodAge), CInt(sCodExpe))
        oFeriado = RecuperaFeriadosEDPYME(dFecha)

        'nMontoGasto = ObtieneGastoTotalCal(nPrestamo, 1, CInt(D.ObtenerDato("nProd")), CInt(D.ObtenerDato("nSubProd")))
        nTasaDesgravamen = RecuperaTasaSeguroDesgravamenEDPYME(CInt(sCodExpe), CInt(sCodAge), _
                                    nPrestamo, 1)
        '----------
        MatCalenCuotasDesemParDetalle = New ComTipos.clsDataset
        'MatCalenCuotasDesemParDetalle = Nothing

        If dsCredDesPar.ObtenerDato("nTipoDesemb") = "2" Then
            nPrestamoTempo = 0
            For i = dsCalenDesPar.nNumFilas To 1 Step -1
                If i > CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
                    dsCalenDesPar.EliminaRegistro(i)
                Else
                    If CInt(dsCalenDesPar.ObtenerDato("nItem", i)) = CInt(dsCredDesPar.ObtenerDato("nNroPrestamoPend")) Then
                        dsCalenDesPar.AsignarDato(Format(dFecha, "dd/MM/yyyy"), "dFecha", i)
                    End If
                    nPrestamoTempo = nPrestamoTempo + CDbl(dsCalenDesPar.ObtenerDato("nPrestamo", i))
                End If
            Next

            oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamoTempo, nCuotas, nTasa, nDia, dsCredDesPar.ObtenerDato("dFecVig"), nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, False, nTasaDesgravamen, True, dsCalenDesPar, MatCalenCuotasDesemParDetalle)
            'dsCalDesembRenov = RenovarDesembolso(sCodAge, sCodExpe, dFecha, nValorITF, nPrestamo, pnSoloIntereses)
        Else
            oTemp = oCalend.GeneraCalendario(nMontoGasto, oFeriado, nPrestamo, nCuotas, nTasa, nDia, dFecha, nDiasGracia, nTipoGracia, nGastos, nProd, nSubProd, IIf(CInt(dsCamp.ObtenerDato("nCodCamp")) = 0, True, False), nTipoPeriodo, nPlazo, nValorITF, pnSoloIntereses, nTasaDesgravamen)
        End If

        dsCamp = Nothing

        oCalend = Nothing

        Call dsGastos.CargaTabla(, , "DBSOLIDA..CredW_RecuperaGastosAlDesembDet " & nPrestamo & "," & CInt(sCodExpe) & "," & CInt(sCodAge))

        sGastos = dsGastos.DevuelveXML

        dsGastos = Nothing

        'If dsCredDesPar.ObtenerDato("nTipoDesemb") <> "2" Then

        ReDim MatCalendCol(7)
        MatCalendCol(1) = "nNroCuota"
        MatCalendCol(2) = "dFecVenc"
        MatCalendCol(3) = "nCapital"
        MatCalendCol(4) = "nIntComp"
        MatCalendCol(5) = "nIntMor"
        MatCalendCol(6) = "nIGV"
        MatCalendCol(7) = "nIntGra" 'ARCV 30-10-2007

        With dsCalend
            .CrearVacio(MatCalendCol)
            For i = 1 To UBound(oTemp)
                .AddRegistro()
                .Seek(i)
                .AsignarDato(oTemp(i, 2), "nNroCuota")
                .AsignarDato(CDate(oTemp(i, 1)).ToString("yyyyMMdd"), "dFecVenc")
                .AsignarDato(oTemp(i, 4), "nCapital")
                .AsignarDato(oTemp(i, 5), "nIntComp")
                .AsignarDato(0, "nIntMor")
                .AsignarDato(oTemp(i, 6), "nIGV")
                .AsignarDato(oTemp(i, 9), "nIntGra") 'ARCV 30-10-2007
            Next i
        End With

        sCalend = dsCalend.DevuelveXML
        dOFecVenc = CDate(oTemp(UBound(oTemp), 1)) 'Ultima fecha del Calendario

        'Else
        'sCalend = dsCalDesembRenov.DevuelveXML
        'dOFecVenc = CDate(dsCalDesembRenov.ObtenerDato("dFecVenc", dsCalDesembRenov.nNumFilas)) 'Ultima fecha del Calendario
        'End If


        dsCalend = Nothing



        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("DBSOLIDA..Cred_DesembolsarCredito_EDPYME", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        'oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco

        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@XCalend", SqlDbType.Xml).Value = sCalend
        oComm.Parameters.Add("@XGastos", SqlDbType.Xml).Value = sGastos

        oComm.Parameters.Add("@cOPagare", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOPagare").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODocEmi", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODocEmi").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODomEmi", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cODomEmi").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOConyuge", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOConyuge").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODocCon", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODocCon").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODomCon", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cODomCon").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cODpto", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODpto").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOProv", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOProv").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cODist", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cODist").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOLinea", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOLinea").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOMoneda", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOMoneda").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@bODNI", SqlDbType.Bit)
        oComm.Parameters.Item("@bODNI").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 31-10-2007
        oComm.Parameters.Add("@cOCodBCP", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOCodBCP").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOCodBN", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@cOCodBN").Direction = ParameterDirection.Output
        '-------------

        'ARCV 02-11-2007
        oComm.Parameters.Add("@cOGrupo", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOGrupo").Direction = ParameterDirection.Output
        '----------

        'ARCV 07-07-2008
        'oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cPComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque
        oComm.Parameters.Add("@cRUC", SqlDbType.VarChar, 20).Value = sRUC
        oComm.Parameters.Add("@cOZona", SqlDbType.VarChar, 20).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@XCalenDesemPar", SqlDbType.Xml).Value = MatCalenCuotasDesemParDetalle.DevuelveXML


        oComm.ExecuteNonQuery()

        cOPagare = Trim(oComm.Parameters.Item("@cOPagare").Value)
        cODocEmi = Trim(oComm.Parameters.Item("@cODocEmi").Value)
        cODomEmi = Trim(oComm.Parameters.Item("@cODomEmi").Value)
        cOConyuge = Trim(oComm.Parameters.Item("@cOConyuge").Value)
        cODocCon = Trim(oComm.Parameters.Item("@cODocCon").Value)
        cODomCon = Trim(oComm.Parameters.Item("@cODomCon").Value)

        cODpto = Trim(oComm.Parameters.Item("@cODpto").Value)
        cOProv = Trim(oComm.Parameters.Item("@cOProv").Value)
        cODist = Trim(oComm.Parameters.Item("@cODist").Value)
        cOLinea = Trim(oComm.Parameters.Item("@cOLinea").Value)
        cOMoneda = Trim(oComm.Parameters.Item("@cOMoneda").Value)
        bODNI = oComm.Parameters.Item("@bODNI").Value

        'ARCV 31-10-2007
        cOCodBCP = Trim(IIf(IsDBNull(oComm.Parameters.Item("@cOCodBCP").Value), "", oComm.Parameters.Item("@cOCodBCP").Value))
        cOCodBN = Trim(IIf(IsDBNull(oComm.Parameters.Item("@cOCodBN").Value), "", oComm.Parameters.Item("@cOCodBN").Value))
        '-------
        cOGrupo = Trim(oComm.Parameters.Item("@cOGrupo").Value) 'ARCV 02-11-2007

        'ARCV 07-07-2008
        'sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        'sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sPCOMENTC = CStr(oComm.Parameters.Item("@cPComentC").Value)
        sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        sZona = CStr(oComm.Parameters.Item("@cOZona").Value)
        '------
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function RecuperaTasaSeguroDesgravamenEDPYME(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, _
                                        ByVal pnPrestamo As Double, ByVal pnMoneda As Integer, _
                                            Optional ByVal pnProd As Integer = 0, Optional ByVal pnSubProd As Integer = 0) As Double

        Dim R As ComTipos.clsDataset

        R = New ComTipos.clsDataset
        R.CargaTabla(, , "DBSOLIDA..Cred_RecuperaTasaSeguroDesgravamen " & pnCodAge & "," & pnCodCred & "," & pnPrestamo & "," & pnMoneda & "," & pnProd & "," & pnSubProd)

        RecuperaTasaSeguroDesgravamenEDPYME = CDbl(R.ObtenerDato("nValor"))

        R = Nothing
    End Function

    Public Function RecuperaFeriadosEDPYME(ByVal pdFechaDesem As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DBSOLIDA..Cred_DevuelveFeriados '" & Format(pdFechaDesem, "dd/MM/yyyy") & "'")
        RecuperaFeriadosEDPYME = D
        D = Nothing
    End Function

    Public Sub RecuperaCreditoDatosEDPYME(ByRef oCred As ComTipos.clsDataset, _
                    ByVal nCodCred As Integer, _
                    ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBSOLIDA..Cred_RecuperaCreditoDatos " & nCodCred & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Function RecuperaDuplicadoCalendEDPYME(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsCredDupCal
        Dim D As ComTipos.ClsCredDupCal
        D = New ComTipos.ClsCredDupCal
        D.oDato.CargaTabla(, , "DBSOLIDA..Cred_DuplicCalendDatos " & pnCodAge & "," & pnCodCred)
        D.oCalend.CargaTabla(, , "DBSOLIDA..Cred_DuplicCalend " & pnCodAge & "," & pnCodCred)
        D.oCalenDesemb.CargaTabla(, , "DBSOLIDA..Cred_RecuperaCalDesParcialDetalle " & pnCodAge & "," & pnCodCred)
        RecuperaDuplicadoCalendEDPYME = D
        D = Nothing
    End Function

    Public Sub BuscarDesembolsoGrupal_WINEDPYME(ByRef oCred As ComTipos.clsDataset, _
                    ByVal nGrupo As Integer, _
                    ByVal nCiclo As Integer, _
                    ByVal nCodAge As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBSOLIDA..Cred_BuscarDesembolsoGrupal " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub CargaDatosDesembolsoGrupo_WINEDPYME(ByRef oCred As ComTipos.clsDataset, _
                    ByVal nGrupo As Integer, _
                    ByVal nCiclo As Integer, _
                    ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBSOLIDA..Cred_CargaDatosDesembolsoGrupo " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub CalendarioGrupoEDPYME(ByRef oCred As ComTipos.clsDataset, _
                ByVal nGrupo As Integer, _
                ByVal nCiclo As Integer, _
                ByVal nCodAge As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBSOLIDA..CredW_CalendarioGrupo " & nGrupo & "," & nCiclo & "," & nCodAge)
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RecuperarGruposEDPYME(ByVal nCodAge As Integer, _
                            ByRef oDatos As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBSOLIDA..CredW_RecuperaGrupos " & nCodAge)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CargaInterfazDesembGrupoEDPYME(ByVal nCodAge As Integer, _
            ByRef oDatos As ComTipos.clsDataset, _
            ByRef oBancos As ComTipos.clsDataset)
        oCons = New ComConstante.clsConstante
        Call RecuperarGruposEDPYME(nCodAge, oDatos)
        oBancos = oCons.RecuperaConstantes(2000)
        oCons = Nothing
    End Sub

    Public Function BuscaGruposEDPYME(ByVal psNomGrp As String, ByVal pnCodAge As Integer, Optional ByVal pnCodsubProd As Integer = 1) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DBSOLIDA..Cred_RecuperaGrupos '" & psNomGrp & "'," & pnCodAge & "," & pnCodsubProd)
        BuscaGruposEDPYME = D
        D = Nothing
    End Function

    Public Function RecuperaDatosGrupoEDPYME(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsCredGrupo
        Dim C As ComTipos.ClsCredGrupo
        C = New ComTipos.ClsCredGrupo
        C.DatosGrupo.CargaTabla(, , "DBSOLIDA..Cred_RecuperaDatosDeGrupo " & pnCodAge & "," & pnCodCred)
        C.DatosInteg.CargaTabla(, , "DBSOLIDA..Cred_RecuperaIntegrantesDeGrupo " & pnCodAge & "," & pnCodCred)
        RecuperaDatosGrupoEDPYME = C
        C = Nothing
    End Function

    Public Function ObtieneValorITFEDPYME(ByVal nCodAge As Integer) As Double
        Dim oData As New ComTipos.clsDataset
        ObtieneValorITFEDPYME = 0
        oData.CargaTabla(, , "DBSOLIDA..Cred_RecuperaValorITF " & nCodAge)
        oData.Inicio()
        If oData.nNumFilas > 0 Then ObtieneValorITFEDPYME = oData.ObtenerDato("nITF")
        oData = Nothing
    End Function

    Public Function ObtieneValorITFFIDEICOMISO(ByVal nCodAge As Integer) As Double
        Dim oData As New ComTipos.clsDataset
        ObtieneValorITFFIDEICOMISO = 0
        oData.CargaTabla(, , "DBCORE..Cred_RecuperaValorITF " & nCodAge)
        oData.Inicio()
        If oData.nNumFilas > 0 Then ObtieneValorITFFIDEICOMISO = oData.ObtenerDato("nITF")
        oData = Nothing
    End Function

    Public Function GastosAlDesembolso_CodOpe(ByVal nCodAge As Integer, _
                                                ByVal nCodCred As Integer, _
                                                ByVal nCodOpe As Integer) As Double
        Dim oData As New ComTipos.clsDataset
        GastosAlDesembolso_CodOpe = 0
        oData.CargaTabla(, , "Cred_GastosAlDesembolso_CodOpe " & nCodAge & "," & nCodCred & "," & nCodOpe)
        oData.Inicio()
        If oData.nNumFilas > 0 Then GastosAlDesembolso_CodOpe = oData.ObtenerDato("nMontoGasto")
        oData = Nothing
    End Function

    Public Function GastosAlDesembolso_CodOpe_Grupal(ByVal nCodGrupo As Integer, _
                                            ByVal nCiclo As Integer, _
                                            ByVal nCodOpe As Integer) As Double
        Dim oData As New ComTipos.clsDataset
        Dim oCred As New ComTipos.clsDataset

        GastosAlDesembolso_CodOpe_Grupal = 0

        oData.CargaTabla(, , "Cred_RecuperaCreditosParaDesembolsoGrupo " & nCodGrupo & "," & nCiclo)

        If oData.nNumFilas > 0 Then
            oData.Inicio()
            While Not oData.EOF
                oCred.CargaTabla(, , "Cred_GastosAlDesembolso_CodOpe " & CInt(oData.ObtenerDato("nCodAge")) & "," & CInt(oData.ObtenerDato("nCodCred")) & "," & nCodOpe)
                GastosAlDesembolso_CodOpe_Grupal += oCred.ObtenerDato("nMontoGasto")
                oData.Siguiente()
            End While
        End If
        oData = Nothing
        oCred = Nothing
    End Function

    Public Sub CargaDatosDesembolsoGrupo_CodOpe(ByRef oCred As ComTipos.clsDataset, _
                ByVal nGrupo As Integer, _
                ByVal nCiclo As Integer, _
                ByVal nCodAge As Integer, _
                ByVal nCodOpe As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CargaDatosDesembolsoGrupo_CodOpe " & nGrupo & "," & nCiclo & "," & nCodAge & "," & nCodOpe)
        oCred = oData
        oData = Nothing
    End Sub

    Public Function BuscarCreditos_x_Cliente_Varios_TOTAL(ByVal sCliente As String, _
                                            ByVal nEstado As Integer, ByVal pnTipoBusq As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        If pnTipoBusq = 1 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_TOTAL '" & sCliente & "'," & nEstado)
        End If
        If pnTipoBusq = 2 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_DNI_TOTAL '" & sCliente & "'," & nEstado)
        End If
        If pnTipoBusq = 3 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_Codigo_TOTAL '" & sCliente & "'," & nEstado)
        End If

        Return oData
    End Function

    Public Function RecuperaConstantes() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante
        oData = oCons.RecuperaConstantes(3010)
        RecuperaConstantes = oData
        oData = Nothing
        oCons = Nothing
    End Function

    Public Function ActualizaTipoImpresora(ByVal psCodUsu As String, ByVal pnTipoImpre As Integer) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("SIST_ActualizaTipoImpresora ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@nTipoImpre", SqlDbType.Int).Value = pnTipoImpre
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Function

    Public Function RecuperaCredPersonas(ByVal pnCodAge As Integer, _
                                        ByVal pnCodCred As Integer) As ComTipos.clsDataset

        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredW_RecuperaCredPersonas  " & Trim(Str(pnCodAge)) + "," & Trim(Str(pnCodCred)))
        RecuperaCredPersonas = C
    End Function

    Public Sub ActualizaCredPersonas(ByVal pnCodAge As Integer, _
                                    ByVal pnCodCred As Integer, _
                                    ByVal oRelPers As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizaCredPersonas ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@oRelaPers", SqlDbType.Xml).Value = oRelPers.DevuelveXML

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function ValidaRestriccionesCredTitular(ByVal pCodPers As String) As Boolean
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ValidaRestriccionesCredTitular " & pCodPers)
        'If CInt(oData.ObtenerDato("bOK")) = 0 Then
        ValidaRestriccionesCredTitular = oData.ObtenerDato("bOK") 'False
        'Else
        'ValidaRestriccionesCredTitular = True
        'End If
        oData = Nothing
    End Function

    Public Function ValidaCondicionesSimulador(ByVal pnCodModalidad As Integer, _
                                    ByVal pnDestino As Integer, _
                                    ByVal pnSubProd As Integer, _
                                    ByVal pnDiasGracia As Integer, _
                                    ByVal pdFechaDesemb As Date, _
                                    ByVal pdVencCalend As Date, _
                                    ByVal pnNroCuotas As Integer) As String
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ValidaCondicionesSimulador " & pnCodModalidad & "," & pnDestino & "," & pnSubProd & "," & pnDiasGracia & _
                            ",'" & pdFechaDesemb.ToString("yyyyMMdd") & "','" & pdVencCalend.ToString("yyyyMMdd") & "'," & pnNroCuotas)
        'If CInt(oData.ObtenerDato("bOK")) = 0 Then
        ValidaCondicionesSimulador = oData.ObtenerDato("cRespuesta") 'False
        'Else
        'ValidaRestriccionesCredTitular = True
        'End If
        oData = Nothing
    End Function
    'Cpua ... 04/03/2010 Segundo Afinamiento
    '------------------------------------------------------
    Public Function RecuperaDesembolsosParciales(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Devuelve_DesembolsosParciales " & pnCodAge & ", " & pnCodCred)
        RecuperaDesembolsosParciales = C
        C = Nothing
    End Function
    Public Function RegistraDesembolsosParciales(ByVal pnCodAge As Integer, _
                                ByVal pnCodCred As Integer, _
                                ByVal dFecha As Date, _
                                ByVal pnPrestamo As Double, _
                                ByVal nItem As Integer, _
                                ByVal nTipoAct As Integer) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Mant_DesembolsosParciales", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = pnPrestamo
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = nItem
        oComm.Parameters.Add("@nTipo", SqlDbType.VarChar, 1).Value = nTipoAct
        oComm.Parameters.Add("@nResult", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        RegistraDesembolsosParciales = oComm.Parameters("@nResult").Value

        objCon = Nothing
        oComm = Nothing
    End Function
    '------------------------------------------------------
    Public Function Alerta_SaldovsCalendario() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Alerta_SaldovsCalendario ")
        Alerta_SaldovsCalendario = D
        D = Nothing
    End Function

    Public Function Alerta_ExitenciaAnalistaTitulares() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Alerta_ExitenciaAnalistaTitulares ")
        Alerta_ExitenciaAnalistaTitulares = D
        D = Nothing
    End Function

    Public Function Alerta_CalendarioNegativo() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Alerta_CalendarioNegativo ")
        Alerta_CalendarioNegativo = D
        D = Nothing
    End Function

    Public Function Alerta_CreditoSaldoCapitalCero() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Alerta_CreditoSaldoCapitalCero ")
        Alerta_CreditoSaldoCapitalCero = D
        D = Nothing
    End Function

    Public Function Alerta_MontoPagadoMayorMontoAmortizar() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Alerta_MontoPagadoMayorMontoAmortizar ")
        Alerta_MontoPagadoMayorMontoAmortizar = D
        D = Nothing
    End Function
    Public Function AlertaCreditosTasaCero() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Alerta_CredTasaCero ")
        AlertaCreditosTasaCero = D
        D = Nothing
    End Function

    Public Function BuscarExcepcionReglas(ByVal pnCodPers As Integer, _
                                            ByVal pnResultado As Integer) As String
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarExcepcionReglas " & pnCodPers & "," & pnResultado)
        BuscarExcepcionReglas = oData.ObtenerDato("cRegla") 'False
        oData = Nothing
    End Function

    Public Function BuscarListaReglas(ByVal psDocumento As String, _
                                ByVal pnNroOpe As Integer, _
                                ByVal pbExperto As Boolean, _
                                ByVal pbPropuesta As Boolean) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarListaReglas '" & psDocumento & "'," & pnNroOpe & "," & IIf(pbExperto, 1, 0) & "," & IIf(pbPropuesta, 1, 0))
        BuscarListaReglas = oData
        oData = Nothing
    End Function

    Public Sub ActualizarExcepcionRegla(ByVal pnCodigo As Integer, _
                            ByVal psUser As String, _
                            ByVal pdFecReg As Date, _
                            ByVal pdFecCad As Date, _
                            ByVal psComentario As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ActualizarExcepcionRegla", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodigo", SqlDbType.Int).Value = pnCodigo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = psUser
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = pdFecReg
        oComm.Parameters.Add("@dFecCad", SqlDbType.DateTime).Value = pdFecCad
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 100).Value = psComentario
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function ValidaCreditos(ByVal pdFecha As Date, _
                                ByVal psCodUsu As String, _
                                ByVal pnCodAge As Integer, _
                                ByVal pnCodCred As Integer, _
                                ByVal pnCodPers As Integer, _
                                ByVal pnProducto As Integer, _
                                ByVal pnSubProd As Integer, _
                                ByVal pnPrestamo As Double, _
                                ByVal pdFecVig As Date, _
                                ByVal pnDiasGracia As Integer, _
                                ByVal pnCodModalidad As Integer, _
                                ByVal pnMoneda As Integer, _
                                ByVal pnDestino As Integer, _
                                ByVal pnPlazo As Integer, _
                                ByVal pnCuotas As Integer, _
                                ByVal pnCodGrupo As Integer, _
                                ByVal pbPropuesta As Boolean, _
                                ByVal pnCodCamp As Integer, _
                                ByVal pbRefinanciado As Boolean) As String

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        'Dim nRestriccion As Integer
        Dim sRestriccion As String

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ValidaCreditosIndividuales", objCon.DameConexion)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = pdFecha
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nProducto", SqlDbType.Int).Value = pnProducto
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = pnSubProd
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = pnPrestamo
        oComm.Parameters.Add("@dFecVig", SqlDbType.DateTime).Value = pdFecVig
        oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int).Value = pnDiasGracia
        oComm.Parameters.Add("@nCodModalidad", SqlDbType.Int).Value = pnCodModalidad
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = pnMoneda
        oComm.Parameters.Add("@nDestino", SqlDbType.Int).Value = pnDestino
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int).Value = pnPlazo
        oComm.Parameters.Add("@nCuotas", SqlDbType.Int).Value = pnCuotas
        oComm.Parameters.Add("@nCodGrupo", SqlDbType.Int).Value = pnCodGrupo

        'oComm.Parameters.Add("@nRestriccion", SqlDbType.Int) '
        'oComm.Parameters.Item("@nRestriccion").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cRestriccion", SqlDbType.VarChar, 100) '@cRestriccion
        oComm.Parameters.Item("@cRestriccion").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@bPropuesta", SqlDbType.Bit).Value = IIf(pbPropuesta, 1, 0)
        oComm.Parameters.Add("@nCodCamp", SqlDbType.Int).Value = pnCodCamp

        oComm.Parameters.Add("@bRefinanciado", SqlDbType.Bit).Value = IIf(pbRefinanciado, 1, 0)

        oComm.ExecuteNonQuery()

        sRestriccion = oComm.Parameters.Item("@cRestriccion").Value
        'nRestriccion = oComm.Parameters.Item("@nRestriccion").Value
        objCon = Nothing
        oComm = Nothing

        Return sRestriccion 'CStr(nRestriccion)

    End Function

    Public Function BuscaConsultaExperto(ByVal psDocumento As String, _
                                        ByVal pnCodAge As Integer, _
                                        ByVal pdFechaSist As Date) As String
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscaConsultaExperto '" & psDocumento & "'," & pnCodAge & ",'" & pdFechaSist & "'")
        BuscaConsultaExperto = oData.ObtenerDato("bOK") 'False
        oData = Nothing
    End Function

    Public Function ListaConsultasExperto(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ListaConsultasExperto '" & psDocumento & "'")
        ListaConsultasExperto = oData
        oData = Nothing
    End Function

    Public Function ListaReglasFiltro(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ListaReglasFiltro '" & psDocumento & "'")
        ListaReglasFiltro = oData
        oData = Nothing
    End Function

    Public Sub ActualizarCredReglaExcepcion(ByVal pnCodigo As Integer, _
                        ByVal psUser As String, _
                        ByVal pdFecReg As Date, _
                        ByVal psComentario As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ActualizarCredReglaExcepcion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodigo", SqlDbType.Int).Value = pnCodigo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = psUser
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = pdFecReg
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 100).Value = psComentario
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function BuscarReglasNoExceptuadas(ByVal psDocumento As String) As String
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarReglasNoExceptuadas '" & psDocumento & "'")
        BuscarReglasNoExceptuadas = oData.ObtenerDato("cResultado") 'False
        oData = Nothing
    End Function

    Public Sub GrabarPagoCuotas_Porcentual(ByVal sCodExpe As String, _
                          ByVal sCodAge As String, _
                          ByVal sCodAgeTrans As String, _
                          ByVal dFecha As Date, _
                          ByVal nMonto As Double, _
                          ByVal nCodOpe As Integer, _
                          ByVal nCodBanco As Integer, _
                          ByVal nITF As Double, _
                      ByRef nOCAPITAL As Double, _
                      ByRef nOINTERES As Double, _
                      ByRef nOGASTOS As Double, _
                      ByRef nOMORA As Double, _
                      ByRef nOIGV As Double, _
                      ByRef nOSALDO As Double, _
                      ByRef nOPROXCUOTA As Integer, _
                      ByRef nONROTRANS As Integer, _
                      ByVal sUser As String, _
                      ByRef sOAGEDIR As String, _
                      ByRef sOAGETEL As String, _
                      ByRef sODESCOPE As String, _
                      ByRef sPCOMENTB As String, _
                      ByRef sSCOMENTB As String, _
                      Optional ByVal sNumCheque As String = "", _
                      Optional ByRef sProxVenc As String = "", _
                      Optional ByVal nPagoCta As Integer = 0, Optional ByVal poCuotas As ComTipos.clsDataset = Nothing, _
                      Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PagoCuotas_Porcentual", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque
        oComm.Parameters.Add("@oCalenAmortizado", SqlDbType.Xml).Value = poCuotas.DevuelveXML

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output
        '---------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        objCon = Nothing
        oComm = Nothing
    End Sub

    'eduenas modificado 04/10/2012 --> modificado para pase version 23/11/2012
    Public Function RecuperaMontosNotaCreditoCancelTotal_ExtraJud(ByVal pnNivel As Integer, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_NotaCredCancelTotalRecuperaMontos_ExtraJud " & pnCodAge & "," & pnCodCred & "," & pnNivel)
        'D.CargaTabla(, , "Cred_NotaCredCancelTotalRecuperaMontos_ExtraJud_NEW " & pnCodAge & "," & pnCodCred & "," & pnNivel)
        RecuperaMontosNotaCreditoCancelTotal_ExtraJud = D
        D = Nothing

    End Function

    Public Function BuscarListaReglas_Usuario(ByVal psDocumento As String, _
                            ByVal pnNroOpe As Integer, _
                            ByVal pbExperto As Boolean, _
                            ByVal pbPropuesta As Boolean, _
                            ByVal psCodUsu As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarListaReglas_Usuario '" & psDocumento & "'," & pnNroOpe & "," & IIf(pbExperto, 1, 0) & "," & IIf(pbPropuesta, 1, 0) & ",'" & psCodUsu & "'")
        BuscarListaReglas_Usuario = oData
        oData = Nothing
    End Function

    Public Function CargaDatosPropuestaGrupal(ByVal pnCodAge As Integer, ByVal pdFecSis As Date, ByVal pnCodGrp As Integer, ByVal pnCiclo As Integer, _
                                                ByRef poRegCredito As ComTipos.ClsWRegistroCred) As ComTipos.ClsWCredSolGrp
        Dim D As ComTipos.ClsWCredSolGrp
        D = New ComTipos.ClsWCredSolGrp

        'D = NuevoCreditoGrupal(pnCodAge, pdFecSis)
        D.Integrantes.CargaTabla(, , "Cred_RecuperaIntegrantesDeGrupoPropuesta " & pnCodAge & "," & pnCodGrp)
        If D.Integrantes.nNumFilas > 0 Then
            'D.Integrantes.Inicio()
            'Do While Not D.Integrantes.EOF
            '    D.Integrantes.AsignarDato(D.SeparaActividad(D.Integrantes.ObtenerDato("cActividad")), "cActividad")
            '    D.Integrantes.Siguiente()
            'Loop
            D.Integrantes.Seek(1)
            'D.Datos.CargaTabla(, , "CredW_RecuperaDatosAprGrp " & D.Integrantes.ObtenerDato("nCodCred") & "," & pnCodAge)
            poRegCredito = CargaDatosPropuestaAnalista(D.Integrantes.ObtenerDato("nCodCred"), pnCodAge, pdFecSis)
            'D.Actividades.CargaTabla(, , "CredW_RecuperaActividades")
        End If
        CargaDatosPropuestaGrupal = D

    End Function

    Public Function CargaDatosCreditosPreAprob(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pdFecSis As Date) As ComTipos.ClsWRegistroCred
        Dim C As ComTipos.ClsWRegistroCred

        C = New ComTipos.ClsWRegistroCred
        C.Datos.CargaTabla(, , "CredW_RecuperaCredPreAprob  " & Trim(Str(pnCodCred)) & "," & Trim(Str(pnCodAge)))
        C.cboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.CboCampanias.CargaTabla(, , "CredW_RecuperaCampanias '" & Format(pdFecSis, "dd/MM/yyyy") & "'")
        C.CalDesParcial.CargaTabla(, , "Cred_RecuperaCalDesParcial " & Trim(Str(pnCodAge)) & "," & Trim(Str(pnCodCred)))

        C.cboAnalista.CargaTabla(, , "CredW_RecuperaAnalista  " & Trim(Str(pnCodAge))) 'ARCV 10-10-2011

        CargaDatosCreditosPreAprob = C
    End Function

    Public Sub GrabarCredArchivos(ByVal nCodCred As Integer, _
                            ByVal nCodAge As Integer, _
                             ByVal oArchivo1() As Byte, _
                             ByVal oArchivo2() As Byte, _
                             ByVal sUser As String, _
                             ByVal dFecReg As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GrabarCredArchivos", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@oArchivo1", SqlDbType.Image).Value = IIf(oArchivo1 Is Nothing, DBNull.Value, oArchivo1)
        oComm.Parameters.Add("@oArchivo2", SqlDbType.Image).Value = IIf(oArchivo2 Is Nothing, DBNull.Value, oArchivo2)
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function RecuperaCredArchivos(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer) As ComTipos.clsCredArchivos
        Dim C As ComTipos.clsCredArchivos

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        objCon = New ComConecta.clsConecta

        oComm = New SqlCommand("Cred_RecuperaCredArchivos", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        R = oComm.ExecuteReader
        'oComm.ExecuteNonQuery()

        RecuperaCredArchivos = New ComTipos.clsCredArchivos
        While R.Read
            With RecuperaCredArchivos
                .oArchivo1 = IIf(IsDBNull(R("oArchivo1")), Nothing, R("oArchivo1"))
                .oArchivo2 = IIf(IsDBNull(R("oArchivo2")), Nothing, R("oArchivo2"))
                .sUser = R("cUser")
            End With
        End While
        objCon = Nothing
        oComm = Nothing

    End Function

    Public Sub ActualizarCredRefinancBloqueo(ByVal pnCodigo As Integer, _
                            ByVal psUser As String, _
                            ByRef pdFecReg As Date, _
                            ByVal psComentario As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizarCredRefinancBloqueo", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodigo", SqlDbType.Int).Value = pnCodigo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = psUser
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = pdFecReg
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 100).Value = psComentario

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function BuscarRefinancBloqueoNoExceptuadas(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarRefinancBloqueoNoExceptuadas '" & psDocumento & "'")
        BuscarRefinancBloqueoNoExceptuadas = oData
        oData = Nothing
    End Function

    Public Function BuscarListaCredRefinancBloqueo(ByVal psUsuario As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarListaCredRefinancBloqueo '" & psUsuario & "'")
        BuscarListaCredRefinancBloqueo = oData
        oData = Nothing
    End Function

    Public Function RecuperaGestoresChiclayo() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredRecuperaGestoresChiclayo ")
        RecuperaGestoresChiclayo = oData
        oData = Nothing
    End Function

    Public Sub CargarDataGestoresChiclayo(ByVal pnCodGestor As Integer, _
                    ByVal pnDiasAtrasoIni As Integer, _
                     ByVal pnDiasAtrasoFin As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CargaDataGestores", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGestor", SqlDbType.Int).Value = pnCodGestor
        oComm.Parameters.Add("@nDiasAtrasoIni", SqlDbType.Int).Value = pnDiasAtrasoIni
        oComm.Parameters.Add("@nDiasAtrasoFin", SqlDbType.Int).Value = pnDiasAtrasoFin

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RegistraMontoMinPagoCuota(ByVal nCodAge As Integer, ByVal nCodCred As Integer, _
                    ByVal sUser As String, ByVal cComentario As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RegistraMinPagoCuota", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 10).Value = sUser
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 500).Value = cComentario

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Function Recupera_MontoMinimoPrePagoCuotas(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveMontoMinPrePago " & pnCodAge & ", " & pnCodCred)
        Recupera_MontoMinimoPrePagoCuotas = oData
        oData = Nothing
    End Function

    Public Function DevuelveRutaReportes(ByVal pnCodReporte As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveRutaReporte " & pnCodReporte)
        DevuelveRutaReportes = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaPrincipal(ByVal pcDocumento As String, ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaPrincipalCreditos '" & pcDocumento & "', '" & pcCodUsu & "'")
        RecuperaCreditosBandejaPrincipal = oData
        oData = Nothing
    End Function

    Public Function RecuperaDetalleCreditosBandejaPrincipal(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DetalleBandejaPrincipal " & pnCodAge & ", " & pnCodCred)
        RecuperaDetalleCreditosBandejaPrincipal = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaSolicitud(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaSolicitud '" & pcCodUsu & "'")
        RecuperaCreditosBandejaSolicitud = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaSolicitudGrupal(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaSolicitudGrupal '" & pcCodUsu & "'")
        RecuperaCreditosBandejaSolicitudGrupal = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaPropuesta(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaPropuesta '" & pcCodUsu & "'")
        RecuperaCreditosBandejaPropuesta = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaPropuestaGrupal(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaPropuestaGrupal '" & pcCodUsu & "'")
        RecuperaCreditosBandejaPropuestaGrupal = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaAprobacion(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaAprobacion '" & pcCodUsu & "'")
        RecuperaCreditosBandejaAprobacion = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaAprobacionGeneral(ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaAprobacionGeneral " & pnCodAge)
        RecuperaCreditosBandejaAprobacionGeneral = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaAprobacionGrupal(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaAprobacionGrupal '" & pcCodUsu & "'")
        RecuperaCreditosBandejaAprobacionGrupal = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaAutorizacion(ByVal pcCodUsu As String, ByVal pnCodAge As Integer, ByVal pbActivo As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaAutorizacion '" & pcCodUsu & "', " & pnCodAge & ", " & pbActivo)
        RecuperaCreditosBandejaAutorizacion = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaGarantias(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaGarantias '" & pcCodUsu & "'")
        RecuperaCreditosBandejaGarantias = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaGarantiasGrupales(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaGarantiasGrupal '" & pcCodUsu & "'")
        RecuperaCreditosBandejaGarantiasGrupales = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaEvaluacion(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaEvaluacion '" & pcCodUsu & "'")
        RecuperaCreditosBandejaEvaluacion = oData
        oData = Nothing
    End Function
    Public Function RecuperaGrupos(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "PA_RecuperaGruposXAgencia " & nCodAge)
        RecuperaGrupos = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaEvaluacionInivGrupal(ByVal pcCodUsu As String, ByVal pnIDGrupo As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaEvaluacionIndivGrupal '" & pcCodUsu & "', " & pnIDGrupo)
        RecuperaCreditosBandejaEvaluacionInivGrupal = oData
        oData = Nothing
    End Function

    Public Function RecuperaCreditosBandejaEvaluacionGrupal(ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaEvaluacionGrupal '" & pcCodUsu & "'")
        RecuperaCreditosBandejaEvaluacionGrupal = oData
        oData = Nothing
    End Function

    Public Sub ExonerarMoraGastos(ByVal nCodCred As Integer, _
                             ByVal nCodAge As Integer, _
                             ByVal dFecha As Date, _
                             ByVal nNroCuota As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Ope_ActualizaMoraGastos", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nNroCuota", SqlDbType.Int).Value = nNroCuota
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub BuscarDatosCreditoFechaValuta(ByVal sCodExpe As String, _
                                    ByVal sCodAge As String, _
                                    ByRef oTipoCred As ComTipos.clsCredPagoCuotas, _
                                    ByVal dFecha As Date, _
                                    ByVal nEstado As Integer)

        Dim oData As New ComTipos.clsDataset
        oCons = New ComConstante.clsConstante

        oData.CargaTabla(, , "CredW_BuscaExpedienteFechaValuta '" & sCodAge & "','" & sCodExpe & "','" & dFecha & "'," & nEstado)
        'Return oData
        oTipoCred.Datos = oData
        oTipoCred.Bancos = oCons.RecuperaConstantes(2000)
        oData = Nothing
        oCons = Nothing
    End Sub

    Public Sub GrabarPagoCuotasFechaValuta(ByVal sCodExpe As String, _
                          ByVal sCodAge As String, _
                          ByVal sCodAgeTrans As String, _
                          ByVal dFecha As Date, _
                          ByVal nMonto As Double, _
                          ByVal nCodOpe As Integer, _
                          ByVal nCodBanco As Integer, _
                          ByVal nITF As Double, _
                      ByRef nOCAPITAL As Double, _
                      ByRef nOINTERES As Double, _
                      ByRef nOGASTOS As Double, _
                      ByRef nOMORA As Double, _
                      ByRef nOIGV As Double, _
                      ByRef nOSALDO As Double, _
                      ByRef nOPROXCUOTA As Integer, _
                      ByRef nONROTRANS As Integer, _
                      ByVal sUser As String, _
                      ByRef sOAGEDIR As String, _
                      ByRef sOAGETEL As String, _
                      ByRef sODESCOPE As String, _
                      ByRef sPCOMENTB As String, _
                      ByRef sSCOMENTB As String, _
                      Optional ByVal sNumCheque As String = "", _
                      Optional ByRef sProxVenc As String = "", _
                      Optional ByVal nPagoCta As Integer = 0, _
                      Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PagoCuotasFechaValuta", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output
        '---------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function ListarReglasSolicitudExcepcion() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ListarReglasSolicitudExcepcion ")
        ListarReglasSolicitudExcepcion = oData
        oData = Nothing
    End Function

    Public Sub RegistraReglaExcepcion(ByVal nCodFiltro As Integer, _
                         ByVal nRestriccion As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RegistraReglaExcepcion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodFiltro", SqlDbType.Int).Value = nCodFiltro
        oComm.Parameters.Add("@nRestriccion", SqlDbType.Int).Value = nRestriccion
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub
    'mestrada 05/09/2011

    Public Function DevuelveMoneda() As ComTipos.clsDataset

        Dim oDataMoneda As ComTipos.clsDataset
        oDataMoneda = New ComTipos.clsDataset
        oDataMoneda.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2090")
        DevuelveMoneda = oDataMoneda
        oDataMoneda = Nothing
    End Function
    Public Function DevuelveCargo() As ComTipos.clsDataset

        Dim oDataCargo As ComTipos.clsDataset
        oDataCargo = New ComTipos.clsDataset
        oDataCargo.CargaTabla(, , "CargaCargo")
        DevuelveCargo = oDataCargo
        oDataCargo = Nothing
    End Function
    Public Function DevuelveDatosCargo(ByVal nCodAge As Integer, ByVal nTipo As Integer) As ComTipos.clsDataset
        Dim oDataDatosCargo As ComTipos.clsDataset
        oDataDatosCargo = New ComTipos.clsDataset
        oDataDatosCargo.CargaTabla(, , "DevuelveDatosCargo " & nCodAge & "," & nTipo)
        DevuelveDatosCargo = oDataDatosCargo
        oDataDatosCargo = Nothing
    End Function

    Public Function DevuelveTiposCuenta() As ComTipos.clsDataset
        Dim oDataCuenta As ComTipos.clsDataset
        oDataCuenta = New ComTipos.clsDataset
        oDataCuenta.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5811")
        DevuelveTiposCuenta = oDataCuenta
        oDataCuenta = Nothing
    End Function
    Public Function DevuelveBancosChequeGerencia() As ComTipos.clsDataset
        Dim oDatos As ComTipos.clsDataset
        oDatos = New ComTipos.clsDataset
        oDatos.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5810")
        DevuelveBancosChequeGerencia = oDatos
        oDatos = Nothing
    End Function

    Public Sub GuardaDatosHabilitaTesoreria(ByVal poData As ComTipos.clsDataset, ByRef ngraba As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("GrabaHabilitaTesoreria", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pdatos", SqlDbType.Xml)
        oComm.Parameters("@pdatos").Value = poData.DevuelveXML
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = ngraba
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function DevuelveBandejaTesoreria(Optional ByVal dFechaInicio As String = "", Optional ByVal dFechaFin As String = "") As ComTipos.clsDataset
        Dim oDatt As ComTipos.clsDataset
        oDatt = New ComTipos.clsDataset
        oDatt.CargaTabla(, , "RecuperaDemsembolsoctaChequeTesoreria ") ' & dFechaInicio & "','" & dFechaFin & "'")
        DevuelveBandejaTesoreria = oDatt
        oDatt = Nothing
    End Function
    Public Function DevuelveDatosBenef(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim Data As ComTipos.clsDataset
        Data = New ComTipos.clsDataset
        Data.CargaTabla(, , "DevuelveDatosBeneficiario " & nCodAge & "," & nCodCred)
        DevuelveDatosBenef = Data
        Data = Nothing
    End Function

    Public Function DevuelveDatosBenef2(ByVal nCodPers As Integer) As ComTipos.clsDataset
        Dim dat As ComTipos.clsDataset
        dat = New ComTipos.clsDataset
        dat.CargaTabla(, , "DevuelveDatosBeneficiario2 " & nCodPers)
        DevuelveDatosBenef2 = dat
        dat = Nothing
    End Function
    Public Sub GrabaDatosChequeGerenciaAbonoCta(ByVal nNroTransacc As Integer, ByVal dFecha As Date, ByVal nCodOpe As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer, ByVal nCodBanco As Integer, _
                                               ByVal nMonto As Double, ByVal nMoneda As Integer, ByVal cBenef As String, ByVal cDocBenef As String, ByVal cDirBenef As String, ByVal cTelefBenef As String, _
                                               ByVal cEncargado As String, ByVal cCargoEn As String, ByVal cDocEn As String, ByVal cDirEn As String, ByVal nTipoCuenta As Integer, _
                                               ByVal cNumCuenta As String, ByVal cCCI As String, ByVal nTipoCambio As Double, ByVal nCodEmpleado As Integer, ByRef ngraba As Integer, _
                                                ByRef ddFecha As Date, _
                                                ByRef nnMonto As Double, _
                                                ByRef ccBanco As String, _
                                                ByRef ccMoneda As String, _
                                                ByRef ccBenef As String, _
                                                ByRef ccDocBenef As String, _
                                                ByRef ccDirBenef As String, _
                                                ByRef ccTelefBenef As String, _
                                                ByRef ccEncargado As String, _
                                                ByRef ccCargoEn As String, _
                                                ByRef ccDocEn As String, _
                                                ByRef ccDirEn As String, _
                                                ByRef ccNomCuenta As String, _
                                                ByRef ccNumCuenta As String, _
                                                ByRef ccCCI As String, ByRef nTipoCambioSalida As Double, _
                                                ByRef nMontoSalida As Double, ByRef nMonedaSalida As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("GrabaDesembolsoChequeGerencia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTransacc
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@cBenef", SqlDbType.VarChar, 200).Value = cBenef
        oComm.Parameters.Add("@cDocBenef", SqlDbType.VarChar, 15).Value = cDocBenef
        oComm.Parameters.Add("@cDirBenef", SqlDbType.VarChar, 200).Value = cDirBenef

        If cTelefBenef <> "" Then
            oComm.Parameters.Add("@cTelefBenef", SqlDbType.VarChar, 20).Value = cTelefBenef
        End If

        If cEncargado <> "" Then
            oComm.Parameters.Add("@cEncargado", SqlDbType.VarChar, 100).Value = cEncargado
        End If

        If cCargoEn <> "" Then
            oComm.Parameters.Add("@cCargoEn", SqlDbType.VarChar, 100).Value = cCargoEn
        End If

        If cDocEn <> "" Then
            oComm.Parameters.Add("@cDocEn", SqlDbType.VarChar, 15).Value = cDocEn
        End If

        If cDirEn <> "" Then
            oComm.Parameters.Add("@cDirEn", SqlDbType.VarChar, 200).Value = cDirEn
        End If

        If nTipoCuenta <> 0 Then
            oComm.Parameters.Add("@nTipoCuenta", SqlDbType.Int).Value = nTipoCuenta
        End If

        If cNumCuenta <> "" Then
            oComm.Parameters.Add("@cNumCuenta", SqlDbType.VarChar, 50).Value = cNumCuenta
        End If

        If cCCI <> "" Then
            oComm.Parameters.Add("@cCCI", SqlDbType.VarChar, 50).Value = cCCI
        End If

        If nTipoCambio <> CDbl(0) Then
            oComm.Parameters.Add("@nTipoCambio", SqlDbType.Money).Value = nTipoCambio
        End If

        If nCodEmpleado <> 0 Then
            oComm.Parameters.Add("@nCodEmpleado", SqlDbType.VarChar, 50).Value = nCodEmpleado
        End If





        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ddFecha", SqlDbType.DateTime).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nnMonto", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccBanco", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccMoneda", SqlDbType.VarChar, 20).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccBenef", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccDocBenef", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccDirBenef", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccTelefBenef", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccEncargado", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccCargoEn", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccDocEn", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccDirEn", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccNomCuenta", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccNumCuenta", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@ccCCI", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nTipoCambioSalida", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nMontoSalida", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nMonedaSalida", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        ddFecha = CDate(oComm.Parameters.Item("@ddFecha").Value)
        nnMonto = oComm.Parameters.Item("@nnMonto").Value
        ccBanco = oComm.Parameters.Item("@ccBanco").Value
        ccMoneda = oComm.Parameters.Item("@ccMoneda").Value
        ccBenef = oComm.Parameters.Item("@ccBenef").Value
        ccDocBenef = oComm.Parameters.Item("@ccDocBenef").Value
        ccDirBenef = oComm.Parameters.Item("@ccDirBenef").Value
        ccTelefBenef = oComm.Parameters.Item("@ccTelefBenef").Value
        ccEncargado = oComm.Parameters.Item("@ccEncargado").Value
        ccCargoEn = oComm.Parameters.Item("@ccCargoEn").Value
        ccDocEn = oComm.Parameters.Item("@ccDocEn").Value
        ccDirEn = oComm.Parameters.Item("@ccDirEn").Value
        ccNomCuenta = oComm.Parameters.Item("@ccNomCuenta").Value
        ccNumCuenta = oComm.Parameters.Item("@ccNumCuenta").Value
        ccCCI = oComm.Parameters.Item("@ccCCI").Value
        nTipoCambioSalida = oComm.Parameters.Item("@nTipoCambioSalida").Value
        nMontoSalida = oComm.Parameters.Item("@nMontoSalida").Value
        nMonedaSalida = oComm.Parameters.Item("@nMonedaSalida").Value
        oComm = Nothing
        objCon = Nothing

    End Sub




    Public Function RecuperaRepresentanteLegalCredito(ByVal sCodExpe As String, _
                            ByVal sCodAge As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaRepresentanteLegalCredito '" & sCodExpe & "','" & sCodAge & "'")
        RecuperaRepresentanteLegalCredito = oData
        oData = Nothing
    End Function

    Public Function BuscaImpresionDocPista(ByVal nCodAge As Integer, _
                        ByVal nCodCred As Integer, _
                        ByVal nTipo As Integer, _
                        ByVal nCodGrp As Integer, _
                        ByVal nCiclo As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscaImpresionDocPista " & nCodAge & "," & nCodCred & "," & nTipo & "," & nCodGrp & "," & nCiclo)
        BuscaImpresionDocPista = oData
        oData = Nothing
    End Function

    Public Sub GrabarPagoLibreAmortizac(ByVal sCodExpe As String, _
                              ByVal sCodAge As String, _
                              ByVal sCodAgeTrans As String, _
                              ByVal dFecha As Date, _
                              ByVal nMonto As Double, _
                              ByVal nCodOpe As Integer, _
                              ByVal nCodBanco As Integer, _
                              ByVal nITF As Double, _
                          ByRef nOCAPITAL As Double, _
                          ByRef nOINTERES As Double, _
                          ByRef nOGASTOS As Double, _
                          ByRef nOMORA As Double, _
                          ByRef nOIGV As Double, _
                          ByRef nOSALDO As Double, _
                          ByRef nOPROXCUOTA As Integer, _
                          ByRef nONROTRANS As Integer, _
                          ByVal sUser As String, _
                          ByRef sOAGEDIR As String, _
                          ByRef sOAGETEL As String, _
                          ByRef sODESCOPE As String, _
                          ByRef sPCOMENTB As String, _
                          ByRef sSCOMENTB As String, _
                          Optional ByVal sNumCheque As String = "", _
                          Optional ByRef sProxVenc As String = "", _
                          Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        'ARCV 11-05-2011
        'oComm = New SqlClient.SqlCommand("Cred_PagoCuotas_Prepago", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_PagoCuotas_LibreAmortizac")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser 'ARCV 29-10-2007

        'ARCV 29-10-2007
        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output
        '---------------

        'oComm.ExecuteNonQuery()
        'ARCV 11-05-2011
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        objCon = Nothing
        oComm = Nothing
    End Sub

    'Napoleon Silva
    'Campañas Enlatados 12/09/2011
    Public Function RecuperaDetalleCampañaEnlatados(ByVal nCodCampania As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset

        C.CargaTabla(, , "Cred_CampaniasEnlatadoDetalle " & nCodCampania)
        RecuperaDetalleCampañaEnlatados = C

        C = Nothing
    End Function

    '************************************************************************************************************************
    'CREADO POR: CPUA
    'USUARIO CREACION : CPUA
    'FECHA CREACION: 13/12/2011
    'USUARIO MODIFICACION:
    'FECHA DE MODIFICACION:
    'COMENTARIO MODIFICACION:
    'Metodo para Registrar las observaciones de aprobacion en bandeja de autorizacion de creditos grupales
    '************************************************************************************************************************
    Public Sub RegistraObsBandejaGrupal(ByVal pnCodGrupo As Integer, ByVal pnCiclo As Integer, ByVal psObs As String, ByVal psCodUsu As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraObservacionesGrupales ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGrupo", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@nCiclo", SqlDbType.Int).Value = pnCiclo
        oComm.Parameters.Add("@sObs", SqlDbType.VarChar, 350).Value = psObs
        oComm.Parameters.Add("@sCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    'mestrada
    Public Sub GrabarPrePagoCuotasFechaValuta(ByVal sCodExpe As String, _
                           ByVal sCodAge As String, _
                           ByVal sCodAgeTrans As String, _
                           ByVal dFecha As Date, _
                           ByVal nMonto As Double, _
                           ByVal nCodOpe As Integer, _
                           ByVal nCodBanco As Integer, _
                           ByVal nITF As Double, _
                       ByRef nOCAPITAL As Double, _
                       ByRef nOINTERES As Double, _
                       ByRef nOGASTOS As Double, _
                       ByRef nOMORA As Double, _
                       ByRef nOIGV As Double, _
                       ByRef nOSALDO As Double, _
                       ByRef nOPROXCUOTA As Integer, _
                       ByRef nONROTRANS As Integer, _
                       ByVal sUser As String, _
                       ByRef sOAGEDIR As String, _
                       ByRef sOAGETEL As String, _
                       ByRef sODESCOPE As String, _
                       ByRef sPCOMENTB As String, _
                       ByRef sSCOMENTB As String, _
                       ByRef nGraba As Integer, ByRef cGraba As String, _
                       Optional ByVal sNumCheque As String = "", _
                       Optional ByRef sProxVenc As String = "", _
                       Optional ByRef nOSEGURO As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PrePagoCuotas_FechaValuta", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCod_Age", SqlDbType.VarChar).Value = sCodAge
        oComm.Parameters.Add("@cCod_Expe", SqlDbType.VarChar).Value = sCodExpe
        oComm.Parameters.Add("@cCod_AgeTrans", SqlDbType.VarChar).Value = sCodAgeTrans
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@nCodOpe", SqlDbType.Int).Value = nCodOpe
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@nITF", SqlDbType.Money).Value = nITF

        oComm.Parameters.Add("@nOCAPITAL", SqlDbType.Money)
        oComm.Parameters.Item("@nOCAPITAL").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOINTERES", SqlDbType.Money)
        oComm.Parameters.Item("@nOINTERES").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOGASTOS", SqlDbType.Money)
        oComm.Parameters.Item("@nOGASTOS").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOMORA", SqlDbType.Money)
        oComm.Parameters.Item("@nOMORA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOIGV", SqlDbType.Money)
        oComm.Parameters.Item("@nOIGV").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOSALDO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSALDO").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nOPROXCUOTA", SqlDbType.Int)
        oComm.Parameters.Item("@nOPROXCUOTA").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nONROTRANS", SqlDbType.Int)
        oComm.Parameters.Item("@nONROTRANS").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser

        oComm.Parameters.Add("@cOAGEDIR", SqlDbType.VarChar, 100)
        oComm.Parameters.Item("@cOAGEDIR").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cOAGETEL", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cOAGETEL").Direction = ParameterDirection.Output
        '----
        'ARCV 27-03-2007
        oComm.Parameters.Add("@cODESCOPE", SqlDbType.VarChar, 50)
        oComm.Parameters.Item("@cODESCOPE").Direction = ParameterDirection.Output
        '----

        'ARCV 07-07-2008
        oComm.Parameters.Add("@cPComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cSComentB", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cProxVenc", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@cSComentC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
        '--------

        oComm.Parameters.Add("@cNumCheque", SqlDbType.VarChar, 20).Value = sNumCheque

        '---------------
        'oComm.Parameters.Add("@bPagoCuenta", SqlDbType.Bit).Value = nPagoCta
        '-------------

        'ARCV 11-12-2009
        oComm.Parameters.Add("@nOSEGURO", SqlDbType.Money)
        oComm.Parameters.Item("@nOSEGURO").Direction = ParameterDirection.Output

        '------ mestrada 27-10-2011
        oComm.Parameters.Add("@nGraba", SqlDbType.Int)
        oComm.Parameters.Item("@nGraba").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@Mensaje", SqlDbType.VarChar, 20)
        oComm.Parameters.Item("@Mensaje").Direction = ParameterDirection.Output
        '---------------

        oComm.ExecuteNonQuery()

        nOCAPITAL = oComm.Parameters.Item("@nOCAPITAL").Value
        nOINTERES = oComm.Parameters.Item("@nOINTERES").Value
        nOGASTOS = oComm.Parameters.Item("@nOGASTOS").Value
        nOMORA = oComm.Parameters.Item("@nOMORA").Value
        nOIGV = oComm.Parameters.Item("@nOIGV").Value
        nOSALDO = oComm.Parameters.Item("@nOSALDO").Value
        nOPROXCUOTA = oComm.Parameters.Item("@nOPROXCUOTA").Value
        nONROTRANS = oComm.Parameters.Item("@nONROTRANS").Value

        'ARCV 29-10-2007
        sOAGEDIR = oComm.Parameters.Item("@cOAGEDIR").Value
        sOAGETEL = oComm.Parameters.Item("@cOAGETEL").Value
        '----
        sODESCOPE = oComm.Parameters.Item("@cODESCOPE").Value  'ARCV 27-03-2008

        'ARCV 07-07-2008
        sPCOMENTB = CStr(oComm.Parameters.Item("@cPComentB").Value)
        sSCOMENTB = CStr(oComm.Parameters.Item("@cSComentB").Value)
        sProxVenc = CStr(oComm.Parameters.Item("@cProxVenc").Value)
        'sSCOMENTC = CStr(oComm.Parameters.Item("@cSComentC").Value)
        '------

        nOSEGURO = oComm.Parameters.Item("@nOSEGURO").Value 'ARCV 11-12-2009

        'mestrada
        nGraba = oComm.Parameters.Item("@nGraba").Value
        cGraba = oComm.Parameters.Item("@Mensaje").Value
        '''''''''
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function RecuperaDatosUsuarioRegistro(ByVal nTipo As Integer, ByVal cCodUsu As String, ByVal nCodPers As Integer) As ComTipos.clsDataset
        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim DataUsu As ComTipos.clsDataset
        DataUsu = New ComTipos.clsDataset
        DataUsu.CargaTabla(, , "RecuperaDatosUsuarioRegistro " & nTipo & ",'" & cCodUsu & "'," & nCodPers)
        RecuperaDatosUsuarioRegistro = DataUsu
        DataUsu = Nothing
    End Function
    Public Function BuscaEmpleado(ByVal pnTipoBusq As Integer, ByVal psCriterioBusq As String) As ComTipos.ClsTipoBuscaEmp
        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim C As ComTipos.ClsTipoBuscaEmp
        C = New ComTipos.ClsTipoBuscaEmp
        C.oDatosEmpBusca.CargaTabla(, , "Recupera_RRHH_BuscaPersona " & pnTipoBusq & ",'" & psCriterioBusq & "'", 1)
        BuscaEmpleado = C
        C = Nothing

    End Function
    Public Sub GrabaSolicitudRegistroCambio(ByVal pGruposUsuSolicitante As ComTipos.clsDataset, _
                                                ByVal pGruposUsuInterino As ComTipos.clsDataset, _
                                                ByVal cUsuSolicitud As String, ByVal nCodCargoUsuSolicitud As Integer, _
                                                ByVal cCargoUsuSolicitud As String, ByVal nCodPersUsuSolcitudNegocio As Integer, _
                                                ByVal nCodPersUsuSolicitudRRHH As Integer, ByVal cUsuInterino As String, _
                                                ByVal cCargoUsuInterino As String, ByVal cCargoUsuInterinoNormal As String, ByVal nCodPersUsuInterinoNegocio As Integer, _
                                                ByVal nCodPersUsuInterinoRRHH As Integer, ByVal cMotivo As String, ByVal dFechaInicioCambio As Date, _
                                                ByVal dFechaFinCambio As Date, ByVal nCodcargoUsuInterinoo As Integer, ByRef ngraba As Integer)



        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("GrabaSolicitudRegistroCambio", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pGruposUsuSolicitante", SqlDbType.Xml)
        oComm.Parameters("@pGruposUsuSolicitante").Value = pGruposUsuSolicitante.DevuelveXML
        oComm.Parameters.Add("@pGruposUsuInterino", SqlDbType.Xml)
        oComm.Parameters("@pGruposUsuInterino").Value = pGruposUsuInterino.DevuelveXML
        oComm.Parameters.Add("@cUsuSolicitud", SqlDbType.VarChar, 50).Value = cUsuSolicitud
        oComm.Parameters.Add("@nCodCargoUsuSolicitud", SqlDbType.Int).Value = nCodCargoUsuSolicitud
        oComm.Parameters.Add("@cCargoUsuSolicitud", SqlDbType.VarChar, 100).Value = cCargoUsuSolicitud

        oComm.Parameters.Add("@nCodPersUsuSolcitudNegocio", SqlDbType.Int).Value = nCodPersUsuSolcitudNegocio
        oComm.Parameters.Add("@nCodPersUsuSolicitudRRHH", SqlDbType.Int).Value = nCodPersUsuSolicitudRRHH

        oComm.Parameters.Add("@cUsuInterino", SqlDbType.VarChar, 50).Value = cUsuInterino
        oComm.Parameters.Add("@cCargoUsuInterino", SqlDbType.VarChar, 100).Value = cCargoUsuInterino

        oComm.Parameters.Add("@nCodPersUsuInterinoNegocio", SqlDbType.Int).Value = nCodPersUsuInterinoNegocio
        oComm.Parameters.Add("@nCodPersUsuInterinoRRHH", SqlDbType.Int).Value = nCodPersUsuInterinoRRHH

        oComm.Parameters.Add("@cCargoUsuInterinoNormal", SqlDbType.VarChar, 100).Value = cCargoUsuInterinoNormal
        oComm.Parameters.Add("@cMotivo", SqlDbType.Text).Value = cMotivo
        oComm.Parameters.Add("@dFechaInicioCambio", SqlDbType.DateTime).Value = dFechaInicioCambio
        oComm.Parameters.Add("@dFechaFinCambio", SqlDbType.DateTime).Value = dFechaFinCambio

        oComm.Parameters.Add("@nCodcargoUsuInterinoo", SqlDbType.Int).Value = nCodcargoUsuInterinoo

        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = ngraba
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function RecuperaBandejaAutorizacion(ByVal cCodUsu As String, ByVal nEstado As Integer) As ComTipos.clsDataset
        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim RecuperaData As ComTipos.clsDataset
        RecuperaData = New ComTipos.clsDataset
        RecuperaData.CargaTabla(, , "RecuperaBandejaAutorizaciones '" & CStr(cCodUsu) & "'," & nEstado)
        RecuperaBandejaAutorizacion = RecuperaData
        RecuperaData = Nothing
    End Function
    Public Function RecuperaComboEstado() As ComTipos.clsDataset

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011

        Dim RecuperaData As ComTipos.clsDataset
        RecuperaData = New ComTipos.clsDataset
        RecuperaData.CargaTabla(, , "RecuperaComboEstado")
        RecuperaComboEstado = RecuperaData
        RecuperaData = Nothing
    End Function
    Public Sub ActualizaEstadoRegistroSolicitud(ByVal nIdRegistro As Integer, ByVal cCodUsu As String, ByRef ngraba As Integer)

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("ActualizaEstadoRegistroSolicitud", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nIdRegistro", SqlDbType.Int).Value = nIdRegistro
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = ngraba
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing
    End Sub
    Public Sub VerificaRegistroSolicitud(ByVal cCodUsu As String, ByRef cEstado As String, ByRef nEstado As Integer)

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("VerificaRegistroSolicitud", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
        oComm.Parameters.Add("@cEstado", SqlDbType.VarChar, 50).Value = cEstado
        oComm.Parameters.Item("@cEstado").Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = nEstado
        oComm.Parameters.Item("@nEstado").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        nEstado = IIf(True = IsDBNull(oComm.Parameters.Item("@nEstado").Value), -1, oComm.Parameters.Item("@nEstado").Value)
        cEstado = IIf(True = IsDBNull(oComm.Parameters.Item("@cEstado").Value), "", oComm.Parameters.Item("@cEstado").Value)
        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Sub EjecutaProcesoCambioSolicitud(ByVal nIdRegistro As Integer, ByVal cCodUsu As String, ByRef nresultado As Integer)
        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("EjecutaProcesoCambioSolicitud", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nIdRegistro", SqlDbType.Int).Value = nIdRegistro
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = nresultado
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nresultado = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing
    End Sub
    Public Sub AnularRegistroEncargatura(ByVal nIdRegistro As Integer, ByVal cCodUsu As String, ByRef nresultado As Integer)
        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("ActualizaEstadoRegistroSolicitud", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nIdRegistro", SqlDbType.Int).Value = nIdRegistro
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 0
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = nresultado
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nresultado = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaGruposEncargatura(ByVal nIdRegistro As Integer, ByVal nTipo As Integer) As ComTipos.clsDataset

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim RecuperaDataEncarga As ComTipos.clsDataset
        RecuperaDataEncarga = New ComTipos.clsDataset
        RecuperaDataEncarga.CargaTabla(, , "RecuperaGruposEncargatura " & nIdRegistro & "," & nTipo)
        RecuperaGruposEncargatura = RecuperaDataEncarga
        RecuperaDataEncarga = Nothing

    End Function
    Public Function RecuperaDataUsuarioDominio(ByVal nCodTIpo As Integer) As ComTipos.clsDataset

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim RecuperaInfo As ComTipos.clsDataset
        RecuperaInfo = New ComTipos.clsDataset
        RecuperaInfo.CargaTabla(, , "RecuperaDataUsuarioDominio " & nCodTIpo)
        RecuperaDataUsuarioDominio = RecuperaInfo
        RecuperaInfo = Nothing

    End Function
    Public Sub EjecutaProcesoCambioSolicitudCierreDia(ByRef nresultado As Integer, ByRef nEjecutado As Integer, ByRef nFinalizado As Integer)

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("EjecutaProcesoCambioSolicitudCierreDia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = nresultado
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        nresultado = oComm.Parameters.Item("@ngraba").Value

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub ActualizaContraseñaVarNegocio(ByVal cValorVar As String, ByVal cCodUsu As String, ByRef nResultado As Integer)

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("ActualizaContraseñaVarNegocio", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cValorVar", SqlDbType.VarChar, 100).Value = nResultado
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 100).Value = nResultado


        oComm.Parameters.Add("@nresultado", SqlDbType.Int).Value = nResultado
        oComm.Parameters.Item("@nresultado").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nResultado = oComm.Parameters.Item("@nresultado").Value
        oComm = Nothing
        objCon = Nothing



    End Sub
    Public Function RecuperaDatosUsuarioRevocatoria(ByVal cCodUsu As String) As ComTipos.clsDataset

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim RecuperaInfo As ComTipos.clsDataset
        RecuperaInfo = New ComTipos.clsDataset
        RecuperaInfo.CargaTabla(, , "RecuperaDatosUsuarioRevocatoria '" & cCodUsu & "'")
        RecuperaDatosUsuarioRevocatoria = RecuperaInfo
        RecuperaInfo = Nothing
    End Function
    Public Sub RevocatoriaxRegistro(ByVal nIdRegistro As Integer, ByVal cCodUsu As String, ByVal ngraba As Integer)

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("RevocatoriaxRegistro", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nIdRegistro", SqlDbType.Int).Value = nIdRegistro
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = ngraba
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaRegistrosCambioDominio(ByVal nTipo As Integer) As ComTipos.clsDataset

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim DataResultado As ComTipos.clsDataset
        DataResultado = New ComTipos.clsDataset
        DataResultado.CargaTabla(, , "RecuperaRegistrosCambioDominio " & nTipo)
        RecuperaRegistrosCambioDominio = DataResultado
        DataResultado = Nothing
    End Function
    Public Function RecuperaBandejaEjecutaEncargatura()

        'CREADO POR
        'USAURIO CREACION : MESTRADA
        'FECHA CREACION : 05/12/2011
        Dim oDataBandejaEjecuta As ComTipos.clsDataset
        oDataBandejaEjecuta = New ComTipos.clsDataset
        oDataBandejaEjecuta.CargaTabla(, , "RecuperaBandejaEjecutaEncargatura")
        RecuperaBandejaEjecutaEncargatura = oDataBandejaEjecuta
        oDataBandejaEjecuta = Nothing
    End Function

    Public Sub Actualiza_LLamadasCallCenter(ByVal pnCodigo As Integer, ByVal pcComentario As String, ByRef ngraba As Integer)
        'CREADO POR
        'USAURIO CREACION : pdesarrollo01
        'FECHA CREACION : 24-04-2012
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActComentLlamada", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodigo", SqlDbType.Int).Value = pnCodigo
        oComm.Parameters.Add("@cComentario2", SqlDbType.NVarChar, 500).Value = pcComentario
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing
    End Sub

    'modificado : pdesarrollo01
    'Fecha :24-04-2012

    Public Function BuscarCreditos_x_Cliente_Varios_Call(ByVal sCliente As String, _
                                                ByVal nEstado As Integer, ByVal pnTipoBusq As Integer, _
                                                ByVal bExcluyeGrupos As Boolean) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        If pnTipoBusq = 1 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_CALL '" & sCliente & "'," & nEstado & "," & IIf(bExcluyeGrupos, 1, 0))
        End If
        If pnTipoBusq = 2 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_DNI_CALL '" & sCliente & "'," & nEstado & "," & IIf(bExcluyeGrupos, 1, 0))
        End If
        If pnTipoBusq = 3 Then
            oData.CargaTabla(, , "Cred_BuscaExpediente_x_Cliente_Varios_Codigo_CALL '" & sCliente & "'," & nEstado & "," & IIf(bExcluyeGrupos, 1, 0))
        End If
        Return oData
    End Function

    Public Function ValidarCreditosCall(ByVal pnCodage As Integer, ByVal pnCodCred As Integer, _
                                               ByVal pcCodusu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredValidaCredGestor " & pnCodage & "," & pnCodCred & ",'" & pcCodusu & "'")
        Return oData
    End Function

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'VVILLAFRANCA  21.05.2012
    Public Function ValidaDireccion(ByVal pnCodPers As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ValidaDireccionSolic", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ValidaDireccion = oComm.Parameters.Item("@nResultado").Value
        oComm = Nothing
        objCon = Nothing
    End Function
    Public Function ValidaUbigeo(ByVal pnCodPers As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ValidaUbigeoSolic", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ValidaUbigeo = oComm.Parameters.Item("@nResultado").Value
        oComm = Nothing
        objCon = Nothing
    End Function

    Public Function ValidaDireccionNego(ByVal pnCodPers As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ValidaDirNegoSolic", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ValidaDireccionNego = oComm.Parameters.Item("@nResultado").Value
        oComm = Nothing
        objCon = Nothing
    End Function

    Public Function ValidaUbigeoNego(ByVal pnCodPers As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ValidaUbigeoSolic", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ValidaUbigeoNego = oComm.Parameters.Item("@nResultado").Value
        oComm = Nothing
        objCon = Nothing
    End Function
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'vvillafranca
    Public Function CargaDatosCreditosObservados(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_DatosComite  " & Trim(pnCodAge) & "," & Trim(pnCodCred))
        CargaDatosCreditosObservados = C
    End Function


    Public Sub RegitraObservacionCredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pdFecReg As Date, ByVal pcUsuReg As String, ByVal pcObservacion As String, ByVal pcComentario As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_InsertObservacion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pdFecReg", SqlDbType.DateTime).Value = pdFecReg
        oComm.Parameters.Add("@pcUsuReg", SqlDbType.VarChar, 20).Value = pcUsuReg
        oComm.Parameters.Add("@pcObservacion", SqlDbType.VarChar, 50).Value = pcObservacion
        oComm.Parameters.Add("@pcComentario", SqlDbType.VarChar, 400).Value = pcComentario
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function ListaCreditosObservados(ByVal pnCodAge) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredListaCreditosObservados " & Trim(pnCodAge))
        ListaCreditosObservados = C
    End Function

    Public Function ListaObservacionesxCredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredListaDetalleObservados " & Trim(pnCodAge) & "," & Trim(pnCodCred))
        ListaObservacionesxCredito = C
    End Function

    Public Sub LevantaObservacionCredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnItem As Integer, ByVal pdFecAprob As Date, ByVal pcUsuAprob As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_LevantaObservacion", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnitem", SqlDbType.Int).Value = pnItem
        oComm.Parameters.Add("@pdFecAprob", SqlDbType.DateTime).Value = pdFecAprob
        oComm.Parameters.Add("@pcUsuAprob", SqlDbType.VarChar, 20).Value = pcUsuAprob
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function RecuperaAgencias() As ComTipos.clsDataset
        Dim D As New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_Devuelve_AgenciasV ")
        RecuperaAgencias = D
        D = Nothing
    End Function


    Public Function ValidaExistenciaObservaciones(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RevisaObservaciones " & pnCodAge & "," & pnCodCred)
        ValidaExistenciaObservaciones = oData
        oData = Nothing
    End Function
    Public Sub ValidaReprogramacionCreditos(ByVal pcCodUsu As String, ByVal pnCodAge As Integer, _
                ByVal pnCodCred As Integer, ByVal pnTipoReprog As Integer, _
                ByVal bFecVenc As Integer, ByRef psRegla As String, ByRef pnNroReprog As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta

        oComm = New SqlClient.SqlCommand("Cred_ValidaReprogramacionCreditos", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nTipoReprog", SqlDbType.Int).Value = pnTipoReprog
        oComm.Parameters.Add("@bFecVenc", SqlDbType.Int).Value = bFecVenc
        oComm.Parameters.Add("@cRestriccion", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nNroReprog", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        psRegla = oComm.Parameters.Item("@cRestriccion").Value
        pnNroReprog = oComm.Parameters.Item("@nNroReprog").Value

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function Cred_BuscarListaReglasReprogramacion(ByVal pnCodAge As Integer, _
                            ByVal pnCodCred As Integer, ByVal pnNroReprog As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarListaReglasReprogramacion " & pnCodAge & ", " & pnCodCred & ", " & pnNroReprog)
        Cred_BuscarListaReglasReprogramacion = oData
        oData = Nothing

    End Function

    Public Sub RegistraSolicitudReprogramacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, _
    ByVal pnNroReprog As Integer, ByVal pnTipoReprog As Integer, ByVal pcCodUsu As String, _
    ByVal pdRepro As ComTipos.ClsCredReprog)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraSolicitudReprogramacion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroReprog", SqlDbType.Int).Value = pnNroReprog
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nTipoReprog", SqlDbType.Int).Value = pnTipoReprog
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@oCalen", SqlDbType.Xml).Value = pdRepro.Calen.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function Cred_BandejaReprogramacion() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BandejaSolicitudReprogramacion ")
        Cred_BandejaReprogramacion = oData
        oData = Nothing
    End Function

    Public Function RecuperaDatosSolicitudReprogramacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnNroReprog As Integer) As ComTipos.ClsCredReprog
        Dim D As ComTipos.ClsCredReprog
        D = New ComTipos.ClsCredReprog

        D.Datos.CargaTabla(, , "Cred_RecuperaDatosBasicosDECredito  " & pnCodAge & "," & pnCodCred)
        D.Calen.CargaTabla(, , "Cred_RecuperaReprog_Calendario " & pnCodCred & "," & pnCodAge & ", " & pnNroReprog)

        RecuperaDatosSolicitudReprogramacion = D
        D = Nothing

    End Function

    Public Sub RegistraAceptacionReprogramacion(ByVal pnNroReprog As Integer, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pdRepro As ComTipos.clsDataset, ByVal pcCodUsu As String, ByVal pcComentarios As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_AceptaSolicitudReprogramacion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroReprog", SqlDbType.Int).Value = pnNroReprog
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@oCalen", SqlDbType.Xml).Value = pdRepro.DevuelveXML
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 1000).Value = pcComentarios

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Sub RegistraRechazoReprogramacion(ByVal pnNroReprog As Integer, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RechazaSolicitudReprogramacion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroReprog", SqlDbType.Int).Value = pnNroReprog
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function RecuperaFechasReprogramacion(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_FechasReprogramacion " & pnCodAge & ", " & pnCodCred)
        RecuperaFechasReprogramacion = oData
        oData = Nothing
    End Function

    Public Function RecuperaCalendarioReprogramacion(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pnNroCalen As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        'Si no funciona echar la culpa a Paola Villena...30/05/2012
        'Siganle echando la culpa
        oData.CargaTabla(, , "Cred_RecuperaReprog_Calendario_Consulta " & pnCodCred & "," & pnCodAge & ", " & pnNroCalen)
        RecuperaCalendarioReprogramacion = oData
        oData = Nothing
    End Function

    Public Function VerificaExcepcionReprogramacion(ByVal pnNroReprog As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_VerificaReglasReprog " & pnNroReprog)
        VerificaExcepcionReprogramacion = oData
        oData = Nothing
    End Function

    Public Function CargaExcepcionesReprogramacion(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CargaExcepcionesReprogramacion " & pnCodAge & ", " & pnCodCred)
        CargaExcepcionesReprogramacion = oData
        oData = Nothing
    End Function

    Public Function CargaNroReprogramaciones(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CargaNroReprogramaciones " & pnCodAge & ", " & pnCodCred)
        CargaNroReprogramaciones = oData
        oData = Nothing
    End Function

    Public Sub RegistraLevantamientoReglaReprogramacion(ByVal pnNroReprog As Integer, ByVal pcCodUsu As String, ByVal pcComentario As String, ByVal pnCodRegla As Integer) ' añadio pnCodRegla - vvillafranca  15/06/2012
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_LevantaExcpecionReprogramacion ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroReprog", SqlDbType.Int).Value = pnNroReprog
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 1000).Value = pcComentario
        oComm.Parameters.Add("@nCodRegla", SqlDbType.Int).Value = pnCodRegla
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'vvillafranca 28-05-2012

    Public Function CargarAgencias() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_Recupera_Agencias ")
        CargarAgencias = oData
        oData = Nothing
    End Function
    Public Function RecuperaDatosClientesExclusivos() As ComTipos.clsCredCliExclusivos
        Dim oClientesExclusivos As New ComTipos.clsCredCliExclusivos
        oClientesExclusivos.oActPrimaria.CargaTabla(, , "Cred_DevuelveTipoActividad ")
        oClientesExclusivos.oCalifSBS.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2166")
        oClientesExclusivos.oNumEntidades.CargaTabla(, , "Cred_NumAgenciasxCred ")
        oClientesExclusivos.oTipoAsignacion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 64")
        oClientesExclusivos.oGeneros.CargaTabla(, , "Cred_ParamSexo ")
        oClientesExclusivos.oTipOrganizacion.CargaTabla(, , "Cred_ParamTipoPersona ")
        oClientesExclusivos.oEstadoCivil.CargaTabla(, , "Cred_ParamEstCivil ")
        oClientesExclusivos.oMeses.CargaTabla(, , "Cred_ParamMeses ")
        oClientesExclusivos.oCampanias.CargaTabla(, , "Cred_listarCampanias ")
        RecuperaDatosClientesExclusivos = oClientesExclusivos
        oClientesExclusivos = Nothing
    End Function

    Public Function DevuelveSubTipoActividad(ByVal pnTipo As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveSubTipoActividad '" & pnTipo & "'")
        DevuelveSubTipoActividad = oData
    End Function

    Public Function CargarCalificacionesSBS() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2166")
        CargarCalificacionesSBS = oData
        oData = Nothing
    End Function

    Public Function CargarNumEntidades() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_NumAgenciasxCred ")
        CargarNumEntidades = oData
        oData = Nothing
    End Function
    Public Function ListadoClientes() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Listado_Clientes ")
        ListadoClientes = oData
        oData = Nothing
    End Function

    Public Function CargarTipoAsign() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 70")
        CargarTipoAsign = oData
        oData = Nothing
    End Function

    Public Function CargarParametros(ByVal ntipoAsig As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "ListarParametros " & ntipoAsig)
        CargarParametros = oData
        oData = Nothing
    End Function

    Public Function ListarClientesSegmentados(ByVal pnTipoAsign As Integer, ByVal pnTipoBusqueda As Integer, _
    ByVal pnCodAge As Integer, ByVal pcCodZona As String, ByVal pnEstado As Integer, ByVal pnEstado2 As Integer, ByVal pnAsignado As Integer, _
    ByVal pnCodAseAsig As Integer, ByVal pnTipoIngreso As Integer, ByVal pcCodUsu As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ListarClientesFiltrados_Estado " & Trim(pnTipoAsign) & "," & Trim(pnTipoBusqueda) & _
        "," & Trim(pnCodAge) & ",'" & Trim(pcCodZona) & "'," & Trim(pnEstado) & "," & Trim(pnEstado2) & "," & Trim(pnAsignado) & "," & Trim(pnCodAseAsig) _
        & "," & Trim(pnTipoIngreso) & ",'" & Trim(pcCodUsu) & "'")
        ListarClientesSegmentados = oData
        oData = Nothing

    End Function

    Public Function ListaAsesores(ByVal nCodAge As Integer, ByVal ntipo As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_ListaAsesoresV   " & nCodAge & "," & ntipo)
        ListaAsesores = D
        D = Nothing
    End Function
    Public Function DevuelveZonaAgencia(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CRED_DevuelveZonaAge   " & nCodAge)
        DevuelveZonaAgencia = D
        D = Nothing
    End Function

    '''++++++++++++++++++++
    Public Function RecuperaGastosxCuotas(ByVal pnCodCred As Integer, ByVal pnCodAge As Integer, ByVal pnNroCuota As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DetalleCalendarioGastos " & pnCodAge & ", " & pnCodCred & ", " & pnNroCuota)
        RecuperaGastosxCuotas = oData
        oData = Nothing
    End Function

    Public Sub AsignaCreditosAnalista(ByVal oData As ComTipos.clsDataset, ByVal nCodPersAna As Integer, ByVal dFecha As Date)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_AsignaAnalista ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPersAna", SqlDbType.Int).Value = nCodPersAna
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = oData.DevuelveXML
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub
    Public Sub ActualizaSeg_Listado(ByVal ntipoUso As Integer, ByVal oData As ComTipos.clsDataset, ByVal nEstado As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizaSegListado ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = ntipoUso
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = nEstado
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = oData.DevuelveXML
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function CargarBandejaVisitasCliPo(ByVal cCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ListarClientesVisitar '" & cCodUsu & "'")
        CargarBandejaVisitasCliPo = oData
        oData = Nothing
    End Function


    Public Function ListarDatosdeVisita(ByVal pnCodPerCli As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredListarDatosComentario " & pnCodPerCli)
        ListarDatosdeVisita = oData
        oData = Nothing
    End Function

    Public Function ListarVisitasCliPo(ByVal pnCodPerCli As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredListarComentarioXCliPot " & pnCodPerCli)
        ListarVisitasCliPo = oData
        oData = Nothing
    End Function

    Public Sub RegistraVisita(ByVal nCodFiltro As Integer, ByVal nCodPers As Integer, ByVal cCodUsu As String, ByVal dFecReg As Date, ByVal nTipoResult As Integer, ByVal cComentario As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistrarComentario ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@ncodFiltro", SqlDbType.Int).Value = nCodFiltro
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = cCodUsu
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@nTipoResul", SqlDbType.Int).Value = nTipoResult
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 400).Value = cComentario
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub ClienteAutomatico_Excel(ByVal oData As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizaCliAutomatico ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = oData.DevuelveXML
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub InsertaPreAprobadoCliPo(ByVal oData As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_InsertPreAprobados ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = oData.DevuelveXML
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function ValidaDesembolsoCliPo(ByVal pncodFiltro As Integer, ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ValidaDesembolsoCliPo " & pncodFiltro & "," & pnCodPers)
        ValidaDesembolsoCliPo = oData
        oData = Nothing
    End Function

    Public Function CargaResultadoVisita() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 8001")
        CargaResultadoVisita = oData
        oData = Nothing
    End Function

    '***************************************************************************************
    'vvillafranca 19-06-2012 - Transferencia Cartera Excel
    '***************************************************************************************
    Public Function ValidaExcelTransferencia(ByVal nCodAseAnt As Integer, ByVal nCodAseNue As Integer, _
                ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ValidaExcelTransfer " & nCodAseAnt & "," & nCodAseNue & "," & nCodAge & "," & nCodCred)
        ValidaExcelTransferencia = oData
        oData = Nothing
    End Function

    Public Sub InsertaTablaTransferencia(ByVal oData As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_CargarTablatransferencia ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = oData.DevuelveXML
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub
    Public Sub TransferenciaCarteraLote()
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_TransferenciaCarteraEnLoteMismaAgencia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Function ValidaTransferenciaNoExista() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ValidaTransferenciaNoExista ")
        ValidaTransferenciaNoExista = oData
        oData = Nothing
    End Function
    '**************************************************************************************

    '***************************************************************************************
    'vvillafranca 26-06-2012 -  Sorteo FiestasPatrias
    '***************************************************************************************
    Public Function ListadodeCuponesSorteo(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ListadoSorteoFP " & nCodAge)
        ListadodeCuponesSorteo = oData
        oData = Nothing
    End Function

    Public Sub RegistraGanadorSorteo(ByVal nCodSerie As Integer, ByVal nCupon As Integer, ByVal nCodSorteo As Integer, ByVal nCodZona As Integer, ByVal nCodPack As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("CredMarcaGanador ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodSerie", SqlDbType.Int).Value = nCodSerie
        oComm.Parameters.Add("@nCupon", SqlDbType.Int).Value = nCupon
        oComm.Parameters.Add("@nCodSorteo", SqlDbType.Int).Value = nCodSorteo
        oComm.Parameters.Add("@nCodZona", SqlDbType.Int).Value = nCodZona
        oComm.Parameters.Add("@nCodPack", SqlDbType.Int).Value = nCodPack
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function DevueldeDatosGanador(ByVal nCodSerie As Integer, ByVal nCupon As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveDatosGanador " & nCodSerie & "," & nCupon)
        DevueldeDatosGanador = oData
        oData = Nothing
    End Function

    Public Function DevueldeDatosSorteo(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveDatosSorteo " & nCodAge)
        DevueldeDatosSorteo = oData
        oData = Nothing
    End Function

    Public Function DevuelvePuntosporTransf(ByVal nNroTrans As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelvePuntosxTransf " & nNroTrans)
        DevuelvePuntosporTransf = oData
        oData = Nothing
    End Function

    Public Sub InsertaCupon(ByVal nCodAge As Integer, ByVal nCodOfic As Integer, ByVal nCodPers As Integer, ByVal nCuponReg As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraNCupones ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodOfic", SqlDbType.Int).Value = nCodOfic
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@nCuponReg", SqlDbType.Int).Value = nCuponReg
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function ValidaCampania(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredValidaSorteoCamp " & nCodAge & "," & nCodCred)
        ValidaCampania = oData
        oData = Nothing
    End Function

    Public Sub ActualizaDirectorioSorteo(ByVal nCodPers As Integer, ByVal cTelefono As String, ByVal cDireccion As String, ByVal cEmail As String, ByVal nOperador As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("CredRegistraDirSorteo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar, 12).Value = cTelefono
        oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 50).Value = cDireccion
        oComm.Parameters.Add("@cEmail", SqlDbType.VarChar, 30).Value = cEmail
        oComm.Parameters.Add("@nOperadorTel", SqlDbType.Int).Value = nOperador
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function DevuelveDatosDirectorio(ByVal nCodPers As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredDevuelvedatos " & nCodPers)
        DevuelveDatosDirectorio = oData
        oData = Nothing
    End Function

    Public Function DevuelveListadoPremios(ByVal nCodSorteo As Integer, ByVal nCodZona As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ListaPremios_FP " & nCodSorteo & "," & nCodZona)
        DevuelveListadoPremios = oData
        oData = Nothing
    End Function

    Public Function ValidaPremiosFP(ByVal nCodSorteo As Integer, ByVal nCodZona As Integer, ByVal nCodPack As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ValidaPremiosxZona_FP " & nCodSorteo & "," & nCodZona & "," & nCodPack)
        ValidaPremiosFP = oData
        oData = Nothing
    End Function

    Public Sub EliminaCupon(ByVal nCodPers As Integer, ByVal nCuponReg As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_AnulaNCupones ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@nCuponReg", SqlDbType.Int).Value = nCuponReg
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub
    '**************************************************************************************
    Public Function CargaComentariosPropuesta(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaComentariosPropuesta  " & Trim(Str(nCodCred)) & "," & Trim(Str(nCodAge)))
        CargaComentariosPropuesta = oData
        oData = Nothing
    End Function


    '**************************************************************************************
    '13/07/2012
    'Public Function RiesgoVinculacionEmp(ByVal nPrestamo As Integer) As ComTipos.clsDataset
    '    Dim oData As New ComTipos.clsDataset
    '    oData.CargaTabla(, , "Cred_RiesgoVinculacionEmpleado " & nPrestamo)
    '    RiesgoVinculacionEmp = oData
    '    oData = Nothing
    'End Function

    Public Sub RiesgoVinculacionEmp(ByVal nPrestamo As Double, ByRef nResult As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RiesgoVinculacionEmpleado", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = nPrestamo
        oComm.Parameters.Add("@nResult", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nResult = oComm.Parameters.Item("@nResult").Value
        objCon = Nothing
        oComm = Nothing

    End Sub

    '10/08/2012
    Public Function Eval_CredVisitar(ByVal cCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        'oData.CargaTabla(, , "Eval_CredVisitar '" & cCodUsu & "'")
        oData.CargaTabla(, , "Eval_CredVisitar_TEMPO '" & cCodUsu & "'")
        Eval_CredVisitar = oData
        oData = Nothing
    End Function

    '10/08/2012
    Public Function Eval_RecuperarSolicitudVisOblig(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperarSolicitudVisOblig " & nCodAge & "," & nCodCred)
        Eval_RecuperarSolicitudVisOblig = oData
        oData = Nothing
    End Function

    '10/08/2012
    Public Sub Eval_RegistrarSegDestinoVisSol(ByVal nCodAge As Integer, ByVal nCodCred As Integer, _
                                              ByVal cCodUsu As String, ByVal cComent As String, _
                                              ByRef nResult As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_RegistrarSegDestinoVisSol", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = cCodUsu
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar).Value = cComent
        oComm.Parameters.Add("@nResult", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        nResult = oComm.Parameters.Item("@nResult").Value
        objCon = Nothing
        oComm = Nothing

    End Sub

    '13/08/2012
    Public Sub Eval_RegistrarVerInversionActFij(ByVal nCodAge As Integer, ByVal nCodCred As Integer, _
                                                ByVal nInvAct As Integer, ByVal nUbiInvAct As Integer, _
                                                ByVal cUbiInvAct As String, ByVal cComentVer As String, _
                                                ByVal nMoneda As Integer, ByVal nRecProp As Double, _
                                                ByVal nRecPropPor As Integer, ByVal nEdpyme As Double, _
                                                ByVal nEdpymePor As Integer, ByVal nOtros As Double, _
                                                ByVal nOtrosPor As Integer, ByVal nSustVerf As Integer, _
                                                ByVal cComentSustVer As String, ByVal nTipoDest As Integer, _
                                                ByVal nEstDestino As Integer, ByVal imagen() As Byte, _
                                                ByVal nEstActMue As Integer, ByVal nValorActMueb As Double, _
                                                ByVal cComentActMue As String, ByVal cComentActInm As String, _
                                                ByVal cInfRelevante As String, ByVal imagenCroquis() As Byte, _
                                                ByVal dFecDesembolso As Date, ByVal dFecVerificacion As Date, _
                                                ByVal nTipo As Integer, ByRef Result As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_RegistrarVerInversionActFij", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nInvAct", SqlDbType.Int).Value = nInvAct
        oComm.Parameters.Add("@nUbiInvAct", SqlDbType.Int).Value = nUbiInvAct
        oComm.Parameters.Add("@cUbiInvAct", SqlDbType.VarChar).Value = cUbiInvAct
        oComm.Parameters.Add("@cComentVerif", SqlDbType.VarChar).Value = cComentVer
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@nRecProp", SqlDbType.Money).Value = nRecProp
        oComm.Parameters.Add("@nRecPropPorc", SqlDbType.Int).Value = nRecPropPor
        oComm.Parameters.Add("@nEdpyme", SqlDbType.Money).Value = nEdpyme
        oComm.Parameters.Add("@nEdpymePorc", SqlDbType.Int).Value = nEdpymePor
        oComm.Parameters.Add("@nOtros", SqlDbType.Money).Value = nOtros
        oComm.Parameters.Add("@nOtrosPorc", SqlDbType.Int).Value = nOtrosPor
        oComm.Parameters.Add("@nSustVerif", SqlDbType.Int).Value = nSustVerf
        oComm.Parameters.Add("@cComentSustVerif", SqlDbType.VarChar).Value = cComentSustVer
        oComm.Parameters.Add("@nTipDestino", SqlDbType.Int).Value = nTipoDest
        oComm.Parameters.Add("@nEstDestino", SqlDbType.Int).Value = nEstDestino
        oComm.Parameters.Add("@imagen", SqlDbType.Image).Value = IIf(imagen Is Nothing, DBNull.Value, imagen)
        oComm.Parameters.Add("@nEstActMue", SqlDbType.Int).Value = nEstActMue
        oComm.Parameters.Add("@nValorActMueb", SqlDbType.Money).Value = nValorActMueb
        oComm.Parameters.Add("@cComentActMueb", SqlDbType.VarChar).Value = cComentActMue
        oComm.Parameters.Add("@cComentActInm", SqlDbType.VarChar).Value = cComentActInm
        oComm.Parameters.Add("@cInfRelevante", SqlDbType.VarChar).Value = cInfRelevante
        oComm.Parameters.Add("@imagenCroquis", SqlDbType.Image).Value = IIf(imagenCroquis Is Nothing, DBNull.Value, imagenCroquis)
        oComm.Parameters.Add("@dFecDesembolso", SqlDbType.DateTime).Value = dFecDesembolso
        oComm.Parameters.Add("@dFecVerificacion", SqlDbType.DateTime).Value = dFecVerificacion
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = nTipo
        oComm.Parameters.Add("@nResult", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        Result = oComm.Parameters.Item("@nResult").Value
        objCon = Nothing
        oComm = Nothing

    End Sub

    '14/08/2012
    Public Sub Eval_BuscarEvaluacionClienteVisita(ByVal cNombCliente As String, ByRef oDatos As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_BuscarEvaluacionClienteVisita '" & cNombCliente & "'")
        oDatos = oData
        oData = Nothing
    End Sub

    '15/08/2012
    Public Function Eval_ListaDesbloqueoCreditos(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        'oData.CargaTabla(, , "Eval_ListaDesbloqueoCreditos " & nCodAge)
        oData.CargaTabla(, , "Eval_ListaDesbloqueoCreditos_TEMPO " & nCodAge)
        Eval_ListaDesbloqueoCreditos = oData
        oData = Nothing
    End Function

    '15/08/2012
    Public Sub Eval_RegistrarDesbloqueoSistema(ByVal pData As ComTipos.clsDataset, ByVal cCodUsu As String, _
                                               ByRef nGraba As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_RegistrarDesbloqueoSistema", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pDatos", SqlDbType.Xml)
        oComm.Parameters.Item("@pDatos").Value = pData.DevuelveXML
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = cCodUsu
        oComm.Parameters.Add("@nGraba", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nGraba = oComm.Parameters.Item("@nGraba").Value
        objCon = Nothing
        oComm = Nothing

    End Sub

    '16/08/2012
    Public Sub Eval_ValidarCredVencidos(ByVal cCodUsu As String, ByRef nResult As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_ValidarCredVencidos", objCon.DameConexion)
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = cCodUsu
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nResult", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nResult = oComm.Parameters.Item("@nResult").Value
        objCon = Nothing
        oComm = Nothing
    End Sub

    '16/08/2012
    Public Function Eval_ListarCredVencidos(ByVal cCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        'oData.CargaTabla(, , "Eval_ListarCredVencidos '" & cCodUsu & "'")
        oData.CargaTabla(, , "Eval_ListarCredVencidos_TEMPO '" & cCodUsu & "'")
        Eval_ListarCredVencidos = oData
        oData = Nothing
    End Function

    Public Function Cred_RecuperaListado(ByVal dFecSist As Date, ByVal cCodUsu As String) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla("", "", "Cred_RecuperaListado '" & Format(dFecSist, "dd/MM/yyyy") & "','" & cCodUsu & "'")

        Cred_RecuperaListado = C

        C = Nothing
    End Function



    Public Function RecuperaModalidades() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos  4546")
        RecuperaModalidades = D
        D = Nothing
    End Function

    Public Function RecuperaTipoCredito() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1500")
        RecuperaTipoCredito = D
        D = Nothing
    End Function

    Public Function RecuperaTipoProductos() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaSubProducto ")
        RecuperaTipoProductos = D
        D = Nothing
    End Function

    Public Function RecuperaDestinos() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        RecuperaDestinos = D
        D = Nothing
    End Function

    Public Function RecuperaEstadosCredito() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaCatalogoCodigos  1700")
        RecuperaEstadosCredito = D
        D = Nothing
    End Function

    Public Sub Actualizar_Modalidad(ByVal pnCodCred As Integer, _
                ByVal pnCodAge As Integer, _
                ByVal pnCodModalidad As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Ope_ActualizaModalidadCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodModalidad", SqlDbType.Int).Value = pnCodModalidad
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Actualizar_TipoCredito(ByVal pnCodCred As Integer, _
                     ByVal pnCodAge As Integer, _
                     ByVal pnCodProd As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Ope_ActualizaProductoCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nProducto", SqlDbType.Int).Value = pnCodProd
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Actualizar_TipoProducto(ByVal pnCodCred As Integer, _
                   ByVal pnCodAge As Integer, _
                   ByVal pnCodSubProd As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Ope_ActualizaSubProductoCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nSubProducto", SqlDbType.Int).Value = pnCodSubProd
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Actualizar_Destino(ByVal pnCodCred As Integer, _
             ByVal pnCodAge As Integer, _
             ByVal pnCodDestino As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Ope_ActualizaDestinoCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nDestino", SqlDbType.Int).Value = pnCodDestino
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function RecuperaDatosCambioCondiciones(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsAutorizaCred
        Dim D As ComTipos.ClsAutorizaCred
        D = New ComTipos.ClsAutorizaCred
        D.oDatos.CargaTabla(, , "Cred_RecuperaDatosBasicos_CambioCondiciones " & pnCodAge & "," & pnCodCred)
        D.oGridAutor.CargaTabla(, , "Cred_RecuperaAutorizaciones " & pnCodCred & "," & pnCodAge)
        RecuperaDatosCambioCondiciones = D
        D = Nothing

    End Function

    Public Function DataCredito_CalificaRiesgo(ByVal sDocumento As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DataCredito_CalificaRiesgo " & sDocumento)
        DataCredito_CalificaRiesgo = D
        D = Nothing
    End Function

    Public Function DataCredito_Consultas_Lista(ByVal sDocumento As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DataCredito_Consultas_ListaEscorf " & sDocumento)
        DataCredito_Consultas_Lista = D
        D = Nothing
    End Function

    '28/08/2012 eduenas
    Public Function Eval_RecuperarVerifActivoFijo(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperarVerifActivoFijo " & nCodAge & "," & nCodCred)
        Eval_RecuperarVerifActivoFijo = oData
        oData = Nothing
    End Function

    'ARCV 14-09-2012
    Public Sub Canje_DevuelveNumeroPuntosDesemb(ByVal nNroTrans As Integer, _
                                    ByRef nNumPuntos As Integer, _
                                    ByRef dFechaI As Date, _
                                    ByRef dFechaF As Date)
        'Dim oData As New ComTipos.clsDataset
        'oData.CargaTabla(, , "Cred_Canje_DevuelveNumeroPuntosDesemb " & nNroTrans)
        'Canje_DevuelveNumeroPuntosDesemb = oData
        'oData = Nothing
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Canje_DevuelveNumeroPuntosDesemb", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTrans
        oComm.Parameters.Add("@nNumPuntos", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@dFechaI", SqlDbType.DateTime).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@dFechaF", SqlDbType.DateTime).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nNumPuntos = oComm.Parameters.Item("@nNumPuntos").Value
        dFechaI = oComm.Parameters.Item("@dFechaI").Value
        dFechaF = oComm.Parameters.Item("@dFechaF").Value
        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Sub Canje_DevuelveNumeroPuntosPago(ByVal nNroTrans As Integer, _
                                    ByRef nNumPuntos As Integer, _
                                    ByRef dFechaI As Date, _
                                    ByRef dFechaF As Date, _
                                    ByRef nNumPuntosAcum As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Canje_DevuelveNumeroPuntosPago", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNroTrans
        oComm.Parameters.Add("@nNumPuntos", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@dFechaI", SqlDbType.DateTime).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@dFechaF", SqlDbType.DateTime).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nNumPuntosAcum", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nNumPuntos = oComm.Parameters.Item("@nNumPuntos").Value
        dFechaI = oComm.Parameters.Item("@dFechaI").Value
        dFechaF = oComm.Parameters.Item("@dFechaF").Value
        nNumPuntosAcum = oComm.Parameters.Item("@nNumPuntosAcum").Value
        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Function Canje_RecuperaListaClientes(ByVal nCodProm As Integer, _
                            ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Canje_RecuperaListaClientes " & nCodProm & " ," & nCodAge)
        Canje_RecuperaListaClientes = oData
        oData = Nothing
    End Function
    Public Sub Canje_RegistraClientes(ByVal pnCodProm As Integer, _
         ByVal poClientes As ComTipos.clsDataset, _
         ByVal psUsuCanje As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Canje_RegistraClientes", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodProm", SqlDbType.Int).Value = pnCodProm
        oComm.Parameters.Add("@oClientes", SqlDbType.Xml).Value = poClientes.DevuelveXML
        oComm.Parameters.Add("@cUsuCanje", SqlDbType.VarChar, 20).Value = psUsuCanje
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub
    '*********************
    'Grabar Expediente 17-10-2012
    Public Function RecuperaPermisoGrabarExpediente(ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredPermiteGrabarExpediente " & pnCodAge)
        RecuperaPermisoGrabarExpediente = D
        D = Nothing
    End Function

    '------------------------------------------------------------------------

    'eduenas 05/11/2012
    Public Sub Registro_ActUsuario01(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario01", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 05/11/2012
    Public Sub Registro_ActUsuario02(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario02", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 05/11/2012
    Public Sub Registro_ActUsuario03(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario03", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 05/11/2012
    Public Sub Registro_ActUsuario04(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario04", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Sub ScoreEjecutaCalculos(ByVal nCodAge As Integer, ByRef nCodcred As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("PA_ScoreGeneraPuntaje", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@pnCodcred", SqlDbType.Int).Value = nCodcred

        oComm.ExecuteNonQuery()


        objCon = Nothing
        oComm = Nothing



    End Sub

    Public Sub ScoreRecuperaCodCredYAge(ByVal nCodEval As Integer, ByRef nCodAge As Integer,
                                        ByRef nCodCred As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("PA_ScoreRecuperaCodCredYAge", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodCred", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        nCodAge = oComm.Parameters("@nCodAge").Value
        nCodCred = oComm.Parameters("@nCodCred").Value


        objCon = Nothing
        oComm = Nothing



    End Sub
    Public Sub RecuperaScoreResultadoPuntaje(ByVal nCodAge As Integer, ByVal nCodCred As Integer,
                                             ByRef nPuntaje As Double, ByRef cResultado As String,
                                             ByRef nResultado As Integer)


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("PA_RecuperaScoreResumenResultado", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nPuntaje", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cResultado", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        nPuntaje = oComm.Parameters("@nPuntaje").Value
        cResultado = oComm.Parameters("@cResultado").Value
        nResultado = oComm.Parameters("@nResultado").Value


        objCon = Nothing
        oComm = Nothing


    End Sub

    Public Function RecuperaScoreDetalle(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "PA_RecuperaScoreDetalle " & nCodAge & "," & nCodCred)
        RecuperaScoreDetalle = D
        D = Nothing

    End Function


    'eduenas 05/11/2012
    Public Sub Registro_ActUsuario05(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario05", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 09/11/2012
    Public Sub Registro_ActUsuario06(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario06", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 09/11/2012
    Public Sub Registro_ActUsuario07(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario07", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 09/11/2012
    Public Sub Registro_ActUsuario08(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario08", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 09/11/2012
    Public Sub Registro_ActUsuario09(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario09", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    'eduenas 09/11/2012
    Public Sub Registro_ActUsuario10(ByVal cUsuario As String, ByVal cVariable As String, _
                                     ByVal cNombPC As String, ByVal cDirIP As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Registro_ActUsuario10", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFechaHora", SqlDbType.DateTime).Value = dFecHora
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar).Value = cUsuario
        oComm.Parameters.Add("@cVariable", SqlDbType.VarChar).Value = cVariable
        oComm.Parameters.Add("@cNombrePC", SqlDbType.VarChar).Value = cNombPC
        oComm.Parameters.Add("@cDirIP", SqlDbType.VarChar).Value = cDirIP
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub


    '========= vvillafranca  05-11-2012 ========

    'Public Function RecuperaMesesCuotaDoble() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    D.CargaTabla(, , "Cred_RecuperaMesesCuotaDoble ")
    '    RecuperaMesesCuotaDoble = D
    '    D = Nothing
    'End Function

    Public Function RecuperaRedondeoxAgencia(ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaRedondeoxAgencia " & pnCodAge) 'cambiar 
        RecuperaRedondeoxAgencia = D
        D = Nothing
    End Function

    Public Function HabilitaCuotaDoble(ByVal nProd As Integer, ByVal nSubProd As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_HabCuotaDoble " & nProd & " ," & nSubProd)
        HabilitaCuotaDoble = D
        D = Nothing
    End Function


    '===========================================

    'eduenas 13/11/2012 - regenerar asientos
    Public Sub Regenerar_Asientos(ByVal pdFecha As Date, ByVal pdFechaMax As Date)
        'Public Sub Regenerar_Asientos(ByVal pdFecha As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RegenerarAsientos_NEW", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = pdFecha
        oComm.Parameters.Add("@dFechaMAX", SqlDbType.DateTime).Value = pdFechaMax
        oComm.CommandTimeout = 0
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing

    End Sub

    ''eduenas 13/11/2012 - supercuadre
    ''Public Function Super_Cuadre(ByVal pdFecha As Date, ByVal pdFechaMax As Date) As ComTipos.clsDataset
    'Public Function Super_Cuadre() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_SuperCuadreAsientos '" & pdFecha & "','" & pdFechaMax & "'")
    '    D.CargaTabla(, , "Cred_SuperCuadreAsientos_NEW ")
    '    Super_Cuadre = D
    '    D = Nothing
    'End Function

    ''eduenas 13/11/2012 - verificacion migracion de asientos
    ''Public Function VerificaMigracionAsiento(ByVal pdFecha As Date) As ComTipos.clsDataset
    'Public Function VerificaMigracionAsiento() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_VerificaMigracionAsiento '" & pdFecha & "'")
    '    D.CargaTabla(, , "Cred_VerificaMigracionAsiento_NEW ")
    '    VerificaMigracionAsiento = D
    '    D = Nothing
    'End Function

    ''eduenas 13/11/2012 - verifica migracion de asientos
    ''Public Function VerificaMigracionAsientoFideicomiso(ByVal pdFecha As Date) As ComTipos.clsDataset
    'Public Function VerificaMigracionAsientoFideicomiso() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_VerificaMigracionAsientoFideicomiso '" & pdFecha & "'",2)
    '    D.CargaTabla(, , "Cred_VerificaMigracionAsientoFideicomiso_NEW ", 2)
    '    VerificaMigracionAsientoFideicomiso = D
    '    D = Nothing
    'End Function

    ''eduenas 13/11/2012 - cuadre de operaciones edpyme/fidecomiso
    ''Public Function CruceOperacionesEdpymeFideicomiso(ByVal pnAño As Integer, ByVal pnMes As Integer) As clsDataset
    'Public Function CruceOperacionesEdpymeFideicomiso() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_CruceOperaciones " & pnAño & ", " & pnMes)
    '    D.CargaTabla(, , "Cred_CruceOperaciones_NEW ")
    '    CruceOperacionesEdpymeFideicomiso = D
    '    D = Nothing
    'End Function

    ''eduenas 13/11/2012 - cuadre efectivo diferencia soles
    ''Public Function CuadreEfecDifSoles(ByVal pnAño As Integer, ByVal pnMes As Integer) As ComTipos.clsDataset
    'Public Function CuadreEfecDifSoles() As ComTipos.clsDataset

    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_CuadreEfecDifSoles " & pnAño & ", " & pnMes)
    '    D.CargaTabla(, , "Cred_CuadreEfecDifSoles_NEW ")
    '    CuadreEfecDifSoles = D
    '    D = Nothing
    'End Function

    ''eduenas 13/11/2012 - cuadre efectivo diferencia dolares
    ''Public Function CuadreEfecDifDolares(ByVal pnAño As Integer, ByVal pnMes As Integer) As ComTipos.clsDataset
    'Public Function CuadreEfecDifDolares() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_CuadreEfecDifDolares " & pnAño & ", " & pnMes)
    '    D.CargaTabla(, , "Cred_CuadreEfecDifDolares_NEW ")
    '    CuadreEfecDifDolares = D
    '    D = Nothing
    'End Function

    ''eduenas 13/11/2012 - cuadre efectivo contable
    ''Public Function CuadreEfecContableEdpyme(ByVal pnAño As Integer, ByVal pnMes As Integer, ByVal pnMoneda As Integer) As ComTipos.clsDataset
    'Public Function CuadreEfecContableEdpyme() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_CuadreEfecContableEdpyme " & pnAño & ", " & pnMes & ", " & pnMoneda)
    '    D.CargaTabla(, , "Cred_CuadreEfecContableEdpyme_NEW ")
    '    CuadreEfecContableEdpyme = D
    '    D = Nothing
    'End Function

    'eduenas 13/11/2012 - actualiza cuenta contable
    Public Sub ValidaAsientosEdpyme(ByVal pnAño As Integer, ByVal pnMes As Integer)
        'Public Sub ValidaAsientosEdpyme()

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Cred_ValidarAsientosEdpyme", objCon.DameConexion)
        oComm = New SqlCommand("Cred_ValidarAsientosEdpyme_NEW", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nAño", SqlDbType.Int).Value = pnAño
        oComm.Parameters.Add("@nMes", SqlDbType.Int).Value = pnMes
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing

    End Sub

    ''eduenas 13/11/2012 - actualiza cuenta contable
    ''Public Function ValidarAsientosEdpyme_Lista(ByVal pnAnio As Integer, ByVal pnMes As Integer) As ComTipos.clsDataset
    'Public Function ValidarAsientosEdpyme_Lista() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Cred_ValidarAsientosEdpyme_Lista " & pnAnio & "," & pnMes, 1)
    '    D.CargaTabla(, , "Cred_ValidarAsientosEdpyme_Lista_NEW ", 1)
    '    ValidarAsientosEdpyme_Lista = D
    '    D = Nothing
    'End Function

    ''eduenas 13/11/2012 - cuadre operaciones Transito
    ''Public Function CruceOperaciones(ByVal pnAño As Integer, ByVal pnMes As Integer) As ComTipos.clsDataset
    'Public Function CruceOperaciones() As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    'D.CargaTabla(, , "Ope_CuadreOperacionesTransito " & pnAño & ", " & pnMes)
    '    D.CargaTabla(, , "Ope_CuadreOperacionesTransito_NEW ")
    '    CruceOperaciones = D
    '    D = Nothing
    'End Function

    Public Function BuscarListaReglasOblig(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscarListaReglasOblig '" & psDocumento & "'")
        BuscarListaReglasOblig = oData
        oData = Nothing
    End Function

End Class

