'***********************************************************************
'CREADO POR: 
'USUARIO CREACION : 
'FECHA CREACION:

'***********************************************************************
'USUARIO MODIFICACION:CCORDOVA
'FECHA DE MODIFICACION:18/07/2012
'COMENTARIO MODIFICACION:
'SE AGREGO LA OPCION DE DESEMBOLSO CON OFICINA COMPARTIDA BN
'***********************************************************************

Imports System.Data
Imports System.Data.SqlClient

Public Class clsTeleCredito
    Dim oCons As ComConstante.clsConstante    
    Public Function RecuperaDatosTeleCredito(ByVal nCargaDatos As Integer, ByVal bEdpyme As Byte, ByVal nCodAge As Integer) As ComTipos.clsRegistroTeleCredito
        Dim D As ComTipos.clsRegistroTeleCredito
        D = New ComTipos.clsRegistroTeleCredito
        If nCargaDatos = 1 Then
            D.Datos.CargaTabla(, , "RecuperaDatosPlanilla_TeleCredito " & bEdpyme & "," & nCodAge)
        End If
        D.CboBancos.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2000")
        D.CboAgencias.CargaTabla(, , "Recupera_AgenciaBN")
        D.oValidacion.CargaTabla(, , "RecuperaVerificacionFechaPlanillaTeleCredito ")
        RecuperaDatosTeleCredito = D
        D = Nothing
    End Function
    Public Function RecuperaDatosCreditosDesembolsados(ByVal nCargaDatos As Integer, ByVal pnCodAge As Integer, ByVal pnCodPers As Integer, ByVal FecIni As Date, ByVal FecFin As Date) As ComTipos.clsRegistroTeleCredito
        Dim D As ComTipos.clsRegistroTeleCredito
        D = New ComTipos.clsRegistroTeleCredito
        If nCargaDatos = 1 Then
            D.Datos.CargaTabla(, , "DevuelveCreditosDesembolsados " & pnCodAge & "," & pnCodPers & ", '" & FecIni & "', '" & FecFin & "'")
        End If
        D.CboBancos.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2000")
        D.CboAgencias.CargaTabla(, , "CredW_Recupera_Agencias")

        RecuperaDatosCreditosDesembolsados = D
        D = Nothing
    End Function
    Public Function RecuperaDesembolsosAnulados(ByVal nCargaDatos As Integer, ByVal pnCodAge As Integer, ByVal pnCodPers As Integer, ByVal FecIni As Date, ByVal FecFin As Date) As ComTipos.clsRegistroTeleCredito
        Dim D As ComTipos.clsRegistroTeleCredito
        D = New ComTipos.clsRegistroTeleCredito
        If nCargaDatos = 1 Then
            D.Datos.CargaTabla(, , "DevuelveAnulacionCreditosDesembolsados " & pnCodAge & "," & pnCodPers & ", '" & FecIni & "', '" & FecFin & "'")
        End If
        D.CboBancos.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2000")
        D.CboAgencias.CargaTabla(, , "CredW_Recupera_Agencias")

        RecuperaDesembolsosAnulados = D
        D = Nothing
    End Function

    Public Function RecuperaPlanilla_a_Abonar_TeleCredito(ByVal nCodAge As Integer, ByVal bEdpyme As Byte, ByVal nCodBanco As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "RecuperaPlanilla_a_Abonar_TeleCredito " & nCodAge & "," & bEdpyme & ", " & nCodBanco)
        RecuperaPlanilla_a_Abonar_TeleCredito = D
        D = Nothing
    End Function
    Public Function RecuperaPlanilla_a_Exportar_TeleCredito(ByVal nCodAge As Integer, ByVal bEdpyme As Byte, ByVal nCodBanco As Integer) As ComTipos.clsRegistroTeleCredito
        Dim D As ComTipos.clsRegistroTeleCredito
        D = New ComTipos.clsRegistroTeleCredito
        D.oExport_Cab.CargaTabla(, , "RecuperaPlanilla_a_Exportar_TeleCredito " & nCodAge & "," & bEdpyme & ", " & nCodBanco)
        D.oExport_Det.CargaTabla(, , "RecuperaPlanilla_a_Exportar_TeleCredito_Det " & nCodAge & "," & bEdpyme & ", " & nCodBanco)
        RecuperaPlanilla_a_Exportar_TeleCredito = D
        D = Nothing
    End Function

    Public Function RecuperaPlanilla_x_Agencia_TeleCredito(ByVal dFecha As Date, ByVal nCodAge As Integer, ByVal nCodBanco As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "RecuperaPlanilla_x_Agencia_TeleCredito '" & dFecha & "'," & nCodAge & ", " & nCodBanco)
        RecuperaPlanilla_x_Agencia_TeleCredito = D
        Call ActualizaEstadoCartasTeleCredito(dFecha, nCodBanco)
        D = Nothing
    End Function
    Public Function ValidaFechaPlanillaTeleCredito() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset

        D.CargaTabla(, , "RecuperaVerificacionFechaPlanillaTeleCredito ")

        ValidaFechaPlanillaTeleCredito = D
        D = Nothing
    End Function

    Public Function VerificaNroPlanillaTeleCredito(ByVal dFecha As Date, ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset

        D.CargaTabla(, , "VerificaNroPlanillaTeleCredito '" & dFecha & "'," & nCodAge)

        VerificaNroPlanillaTeleCredito = D
        D = Nothing
    End Function


    Public Function VerificaHoraMaxima(ByVal nCodAge As Integer, ByVal nCodBanco As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset

        D.CargaTabla(, , "DevuelveTeleCred_HoraMax " & nCodAge & "," & nCodBanco)

        VerificaHoraMaxima = D
        D = Nothing
    End Function
    Public Sub ActualizaEstadoCartasTeleCredito(ByVal dFecha As Date, ByVal nCodBanco As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ActualizaEstadoCartas", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco

        oComm.ExecuteNonQuery()
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub AbonarPlanillaTeleCredito(ByVal oAbonos As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("AbonarPlanillaTeleCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@dAbonos", SqlDbType.Xml).Value = IIf(oAbonos.DevuelveXML = "@#err", "", oAbonos.DevuelveXML)


        oComm.ExecuteNonQuery()
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Function AnularDesembolsoTeleCredito(ByVal cCodUsu As String, ByVal oAbonos As ComTipos.clsDataset) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("RegistraExcepcionesTeleCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@oDesemb", SqlDbType.Xml).Value = IIf(oAbonos.DevuelveXML = "@#err", "", oAbonos.DevuelveXML)
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
        oComm.Parameters.Add("@nResult", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        AnularDesembolsoTeleCredito = oComm.Parameters("@nResult").Value
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Function

    Public Sub ActualizarDesembolsoTeleCredito(ByVal oAbonos As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("ActualizaAnulacionesDesembolsos", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@oDesemb", SqlDbType.Xml).Value = IIf(oAbonos.DevuelveXML = "@#err", "", oAbonos.DevuelveXML)
        oComm.ExecuteNonQuery()
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RegistraPlanillaTeleCredito(ByVal nTipoAct As Integer, _
                                            ByVal dFecha As Date, ByVal bEdpyme As Integer, _
                                            ByVal nCodAge As Integer, ByVal nCodCred As Integer, _
                                            ByVal nPrestamo As Double, ByVal bDesembolsado As Integer, _
                                            ByVal dFechaAbono As Date, ByVal nCodBanco As Integer, _
                                            ByVal cCodUsuReg As String, ByVal cCodUsuProc As String, _
                                            ByVal cNroPlanilla As String, ByVal cNroCarta As String, _
                                            ByVal nCodAgenciaBN As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GrabaPlanillaTeleCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@bEdpyme", SqlDbType.Bit).Value = bEdpyme
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = nPrestamo
        oComm.Parameters.Add("@bDesembolsado", SqlDbType.Bit).Value = bDesembolsado
        oComm.Parameters.Add("@dFechaAbono", SqlDbType.DateTime).Value = dFechaAbono
        oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco
        oComm.Parameters.Add("@cCodUsuReg", SqlDbType.VarChar, 50).Value = cCodUsuReg
        oComm.Parameters.Add("@cCodUsuProc", SqlDbType.VarChar, 50).Value = cCodUsuProc
        oComm.Parameters.Add("@cNroPlanilla", SqlDbType.VarChar, 50).Value = cNroPlanilla
        oComm.Parameters.Add("@cNroCarta", SqlDbType.VarChar, 50).Value = cNroCarta
        oComm.Parameters.Add("@nCodAgenciaBN", SqlDbType.Int).Value = nCodAgenciaBN
        oComm.Parameters.Add("@nTipoAct", SqlDbType.Int).Value = nTipoAct

        oComm.ExecuteNonQuery()
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub SeguridadPlanillaTeleCredito(ByVal nTipoAct As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("SeguridadPlanillaTeleCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nTipoAct", SqlDbType.Int).Value = nTipoAct

        oComm.ExecuteNonQuery()
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function ActualizaNroPlanilla_y_CartaTeleCredito(ByVal cNroPlanilla As String, _
                                                        ByVal dFecha As Date, _
                                                        ByVal nCodAge As Integer) As String

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("ActualizaNroPlanillayCartaTeleCredito", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cNroPlanilla", SqlDbType.VarChar, 50).Value = cNroPlanilla
        oComm.Parameters.Add("@dfecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@cNroCarta", SqlDbType.VarChar, 20).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        objCon.CierraConexion()

        ActualizaNroPlanilla_y_CartaTeleCredito = oComm.Parameters(3).Value

        oComm = Nothing
        objCon = Nothing

    End Function
    Public Sub BuscarBancos(ByRef oDesemb As ComTipos.clsCredDesembolso)
        oCons = New ComConstante.clsConstante        
        oDesemb.Bancos = oCons.RecuperaConstantes(2000)
        oCons = Nothing
    End Sub

    Public Function RecuperaTramaInterbank(ByVal oData As ComTipos.clsDataset) As ComTipos.clsDataset
        Dim oDataTrama As ComTipos.clsDataset
        oDataTrama = New ComTipos.clsDataset
        oDataTrama.CargaTabla(, , "RecuperaDataTrama '" & oData.DevuelveXML & "'")
        RecuperaTramaInterbank = oDataTrama
        oDataTrama = Nothing
    End Function

    Public Sub AbonarPlanillaTeleCreditoInterBank(ByVal dAbonosInterBank As ComTipos.clsDataset, ByRef ngraba As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("AbonarPlanillaTeleCreditoInterBank", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dAbonosInterBank", SqlDbType.Xml).Value = IIf(dAbonosInterBank.DevuelveXML = "@#err", "", dAbonosInterBank.DevuelveXML)
        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        ngraba = oComm.Parameters.Item("@ngraba").Value
        objCon.CierraConexion()

        objCon = Nothing
        oComm = Nothing
    End Sub

    'ARCV 18-07-2012
    Public Function RecuperaPlanilla_a_Abonar_BNOfComp() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "RecuperaPlanilla_a_Abonar_BNOfComp ")
        RecuperaPlanilla_a_Abonar_BNOfComp = D
        D = Nothing
    End Function

    Public Sub AbonarPlanilla_BNOfComp(ByVal oAbonos As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("AbonarPlanilla_BNOfComp", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@dAbonos", SqlDbType.Xml).Value = IIf(oAbonos.DevuelveXML = "@#err", "", oAbonos.DevuelveXML)


        oComm.ExecuteNonQuery()
        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub
End Class
