Imports System.Data
Imports System.Data.SqlClient

Public Class clsEvaluacion
    Public Sub Prejudicial_ReasignaCreditosGestor(ByVal pnXml As ComTipos.clsDataset, _
ByVal pnCodAge As Integer, ByVal pnCodGesorigen As Integer, ByVal nCodPersGestorNew As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Prejudicial_ReasignaCreditosGestor2", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnXml", SqlDbType.Xml).Value = pnXml.DevuelveXML
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodGesOrig", SqlDbType.Int).Value = pnCodGesorigen
        oComm.Parameters.Add("@nCodPersGestorNew", SqlDbType.Int).Value = nCodPersGestorNew

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub
    Public Sub RecuperaGestorCobranzasNegociaciones(ByRef oDatos As ComTipos.clsDataset, _
ByVal nCodAge As Integer, ByVal nCodCred As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Prejudicial_RecuperaGestorCobranzasNegociaciones " & nCodAge & "," & nCodCred)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub Prejudicial_RegistraCalendarioNegociaciones(ByVal pnCabecera As ComTipos.clsDataset, ByVal pnDetalle As ComTipos.clsDataset, _
    ByVal pnCodAge As Integer, ByVal pnCodCred As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Prejudicial_RegistraCalendarioNegociaciones2", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@pnCabecera", SqlDbType.Xml).Value = pnCabecera.DevuelveXML
        oComm.Parameters.Add("@pnDetalle", SqlDbType.Xml).Value = pnDetalle.DevuelveXML
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Sub RecuperaDatosClientePrejudicial(ByRef oDatos As ComTipos.clsDataset, _
   ByVal nCodAge As Integer, ByVal nCodCred As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Prejudicial_BuscarDatosCliente " & nCodAge & "," & nCodCred)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub RecuperaDatosCreditoAgricola(ByRef oDatos As ComTipos.clsDataset, _
           ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CredAgri_RecuperaDatosAgricola " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CargaEvalIndicadoresCreditoAgricola(ByRef oDatos As ComTipos.clsDataset, _
    ByVal nCodEval As Integer)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CredAgri_RecuperaEvalIndCreditoAgricola " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub PersEvalIndCreditoAgricola_Grabar(ByVal nCodEval As Integer, _
                                   ByVal nEndeudamiento As Double, _
                                   ByVal nRentVentas As Double, _
                                   ByVal nCuotaCapPago As Double, _
                                   ByVal nCodCredGarant As Double, _
                                   ByVal nPrestamo As Double, _
                                   ByVal nCuotas As Double, _
                                   ByVal nTasa As Double, _
                                   ByVal nMontoCuota As Double, ByVal nProd As Integer, ByVal nSubProd As Integer, _
                                   ByVal nCodAge As Integer, ByVal nCodCred As Integer, ByVal nDiaFijo As Integer, _
                                   ByVal dFecDes As Date, ByVal nTasaComp As Double, ByVal nDiasGracia As Integer, _
                                   ByVal nGastos As Double, ByVal nCodCamp As Integer, ByVal nTipoPeriodo As Integer, _
                                   ByVal nPlazo As Integer, ByVal pnItfValor As Double, ByVal nSoloIntereses As String, _
                                   ByVal dFecPrimPago As Date, ByVal dCalPar As ComTipos.clsDataset, ByVal nTipoSegDesg As Integer, _
                                   ByVal nTipoGracia As Integer, ByVal nCodLinea As Integer)
        'ByRef nBand As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset


        nMontoGasto = ObtieneGastoTotalCal(CDbl(nPrestamo), 1, CInt(nProd), CInt(nSubProd))
        oFeriado = RecuperaFeriados(CDate(dFecDes))

        oCal = New ComTipos.ClsCalendario

        oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, CInt(nCodAge), CInt(IIf(nCodCred = Nothing, "0", nCodCred)), _
                    1, CDbl(nPrestamo), CInt(nCuotas), CDbl(nTasaComp), _
                    CInt(nDiaFijo), CDate(dFecDes), CInt(nDiasGracia), 0, CDbl(nGastos), _
                    CInt(nProd), CInt(nSubProd), IIf(CInt(nCodCamp) = 0, True, False), CInt(nTipoPeriodo), CInt(nPlazo), pnItfValor, _
                    IIf(nSoloIntereses = "1", True, False), , CDate(dFecPrimPago))



        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_CredAgri_RegistraIndCreditoAgricola", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@pnEndeudamiento", SqlDbType.Money).Value = nEndeudamiento
        oComm.Parameters.Add("@pnRentVentas", SqlDbType.Money).Value = nRentVentas
        oComm.Parameters.Add("@pnCuotaCapPago ", SqlDbType.Money).Value = nCuotaCapPago
        oComm.Parameters.Add("@pnCodCredGarant", SqlDbType.Money).Value = nCodCredGarant
        oComm.Parameters.Add("@pnPrestamo", SqlDbType.Money).Value = nPrestamo
        oComm.Parameters.Add("@pnCuotas", SqlDbType.Int).Value = nCuotas
        oComm.Parameters.Add("@pnTasa ", SqlDbType.Money).Value = nTasa
        oComm.Parameters.Add("@pnMontoCuota ", SqlDbType.Money).Value = nMontoCuota
        oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
        oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML


        oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int).Value = nDiaFijo
        oComm.Parameters.Add("@nCuotas", SqlDbType.Int).Value = nCuotas
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int).Value = nPlazo
        oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int).Value = nDiasGracia
        oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int).Value = nTipoGracia
        oComm.Parameters.Add("@nGastos", SqlDbType.Money).Value = nGastos
        oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt).Value = nTipoPeriodo
        oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.SmallInt).Value = nSoloIntereses
        oComm.Parameters.Add("@nTipoSegDesg", SqlDbType.SmallInt).Value = nTipoSegDesg
        oComm.Parameters.Add("@dFecPrimPago", SqlDbType.DateTime).Value = dFecPrimPago
        oComm.Parameters.Add("@Cal_Des_Par", SqlDbType.Xml).Value = IIf(dCalPar.DevuelveXML = "@#err", "", dCalPar.DevuelveXML)
        oComm.Parameters.Add("@nCodLinea", SqlDbType.SmallInt).Value = nCodLinea



        'oComm.Parameters.Add("@nBand", SqlDbType.Int)
        'oComm.Parameters.Item("@nBand").Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        'nBand = CInt(Trim(oComm.Parameters.Item(5).Value))

        oComm = Nothing
        objCon = Nothing
    End Sub

    'Public Sub PersEvalIndCreditoAgricola_Grabar(ByVal nCodEval As Integer, _
    '                               ByVal nEndeudamiento As Double, _
    '                               ByVal nRentVentas As Double, _
    '                               ByVal nCuotaCapPago As Double, _
    '                               ByVal nCodCredGarant As Double, _
    '                               ByVal nPrestamo As Double, _
    '                               ByVal nCuotas As Double, _
    '                               ByVal nTasa As Double, _
    '                               ByVal nMontoCuota As Double, _
    '                               Optional ByVal DCal As ComTipos.clsDataset = Nothing, _
    '                               Optional ByVal nProducto As Integer = 0, _
    '                                Optional ByVal nSubProducto As Integer = 0, _
    '                                Optional ByVal pnItfValor As Double = 0, _
    '                                Optional ByVal pnCodCamp As Integer = 0, _
    '                                Optional ByVal pdFecDes As Date = #1/1/2000#, _
    '                                Optional ByVal pnCodCred As Integer = 0, _
    '                                Optional ByVal pnCodAge As Integer = 0, _
    '                                Optional ByVal pnNumcuotas As Integer = 0, _
    '                                Optional ByVal pnTasaCom As Double = 0, _
    '                                Optional ByVal pnDiaVenc As Integer = 0, _
    '                                Optional ByVal pnDiasGracia As Integer = 0, _
    '                                Optional ByVal pnTipoGracia As Integer = 0, _
    '                                Optional ByVal pnTipoPer As Integer = 0, _
    '                                Optional ByVal pnPlazo As Integer = 0, _
    '                                Optional ByVal pnCodLin As Integer = 0)
    '    'ByRef nBand As Integer)

    '    Dim objCon As ComConecta.clsConecta
    '    Dim oComm As SqlCommand

    '    Dim oCal As ComTipos.ClsCalendario
    '    Dim oCalData As ComTipos.clsDataset
    '    Dim nMontoGasto As Double
    '    Dim oFeriado As ComTipos.clsDataset
    '    Dim nTasaDesgravamen As Double

    '    nMontoGasto = ObtieneGastoTotalCal(nPrestamo, 1, nProducto, nSubProducto)
    '    oFeriado = RecuperaFeriados(pdFecDes)

    '    'ARCV 21-05-2009
    '    nTasaDesgravamen = RecuperaTasaSeguroDesgravamen(pnCodCred, pnCodAge, _
    '                                nPrestamo, 1, nProducto, nSubProducto)

    '    oCal = New ComTipos.ClsCalendario
    '    oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, pnCodAge, pnCodCred, _
    '                    1, nPrestamo, pnNumcuotas, pnTasaCom, _
    '                    pnDiaVenc, pdFecDes, pnDiasGracia, pnTipoGracia, 0, _
    '                    nProducto, nSubProducto, IIf(pnCodCamp = 0, True, False), pnTipoPer, pnPlazo, pnItfValor, , _
    '                     nTasaDesgravamen)

    '    objCon = New ComConecta.clsConecta
    '    oComm = New SqlCommand("Eval_CredAgri_RegistraIndCreditoAgricola", objCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = nCodEval
    '    oComm.Parameters.Add("@pnEndeudamiento", SqlDbType.Money).Value = nEndeudamiento
    '    oComm.Parameters.Add("@pnRentVentas", SqlDbType.Money).Value = nRentVentas
    '    oComm.Parameters.Add("@pnCuotaCapPago ", SqlDbType.Money).Value = nCuotaCapPago
    '    oComm.Parameters.Add("@pnCodCredGarant", SqlDbType.Money).Value = nCodCredGarant
    '    oComm.Parameters.Add("@pnPrestamo", SqlDbType.Money).Value = nPrestamo
    '    oComm.Parameters.Add("@pnCuotas", SqlDbType.Int).Value = nCuotas
    '    oComm.Parameters.Add("@pnTasa ", SqlDbType.Money).Value = nTasa
    '    oComm.Parameters.Add("@pnMontoCuota", SqlDbType.Money).Value = nMontoCuota

    '    oComm.Parameters.Add("@dCalPar", SqlDbType.Xml).Value = DCal.DevuelveXML
    '    oComm.Parameters.Add("@pdFecDes", SqlDbType.Date).Value = pdFecDes
    '    oComm.Parameters.Add("@pnNumcuotas", SqlDbType.Int).Value = pnNumcuotas
    '    oComm.Parameters.Add("@pnDiaVenc", SqlDbType.Int).Value = pnDiaVenc
    '    oComm.Parameters.Add("@Cal_XML ", SqlDbType.Xml).Value = oCalData.DevuelveXML
    '    oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
    '    oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
    '    oComm.Parameters.Add("@pnCodLin", SqlDbType.Int).Value = pnCodLin
    '    oComm.Parameters.Add("@pnTasaCom ", SqlDbType.Money).Value = pnTasaCom
    '    oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = nPrestamo
    '    oComm.Parameters.Add("@pnTipoPer", SqlDbType.Int).Value = pnTipoPer
    '    oComm.Parameters.Add("@pnPlazo", SqlDbType.Int).Value = pnPlazo
    '    oComm.Parameters.Add("@pnCodLin", SqlDbType.Int).Value = pnCodLin
    '    oComm.Parameters.Add("@pnDiasGracia", SqlDbType.Int).Value = pnDiasGracia
    '    oComm.Parameters.Add("@pnTipoGracia", SqlDbType.Int).Value = pnTipoGracia

    '    'oComm.Parameters.Add("@nBand", SqlDbType.Int)
    '    'oComm.Parameters.Item("@nBand").Direction = ParameterDirection.Output

    '    oComm.ExecuteNonQuery()

    '    'nBand = CInt(Trim(oComm.Parameters.Item(5).Value))

    '    oComm = Nothing
    '    objCon = Nothing
    'End Sub
    Public Sub CargaIndicadoresCreditoAgricola(ByRef oDatos As ComTipos.clsDataset, _
                    ByVal nCodEval As Integer, _
                    ByVal nCodAge As Integer, _
                    ByVal nCodCred As Integer, _
                    ByVal nPrestamo As Double, _
                    ByVal nCuota As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CredAgri_CalculaMontos2 " & nCodEval & "," & nCodAge & "," & nCodCred & "," & nPrestamo & "," & nCuota)
        oDatos = oData
        oData = Nothing
    End Sub


    Public Function RecuperaSupuestosFlujoCaja(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RceuperaSuspuestosFC " & pnCodEval)
        RecuperaSupuestosFlujoCaja = oData
        oData = Nothing

    End Function

    Public Sub BuscarEvaluacion_x_Cliente(ByRef oDatos As ComTipos.clsDataset, _
            ByVal sNomCliente As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_BuscarEvaluacion_x_Cliente '" & sNomCliente & "'")
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CargarEvaluacion(ByRef oDatos As ComTipos.clsDataset, _
            ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CargarEvaluacion " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CargarBotonesEvaluacion(ByRef oDatos As ComTipos.clsDataset, _
            ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CargarBotonesEvaluacion " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub PersEvaluacion_Insertar(ByVal dFecEval As Date, _
                            ByVal nCodPers As Integer, _
                            ByVal sDescripcion As String, _
                            ByVal sNomNegocio As String, _
                            ByVal sDireccion As String, _
                            ByVal sTelefono As String, _
                            ByVal nTipoFteIngreso As Integer, _
                            ByVal sCodZona As String, _
                            ByVal bValidaParamFecha As Integer, _
                            ByRef nError As Integer, _
                            ByRef nCodEval As Integer, _
                            ByVal nCodCred As Integer, _
                            ByVal nCodAge As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Eval_PersEvaluacion_Insertar", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Eval_PersEvaluacion_Insertar")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecEval", SqlDbType.DateTime).Value = dFecEval
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 200).Value = sDescripcion
        oComm.Parameters.Add("@cNomNegocio", SqlDbType.VarChar, 50).Value = sNomNegocio
        oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 50).Value = sDireccion
        oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar, 20).Value = sTelefono
        oComm.Parameters.Add("@nTipoFteIngreso", SqlDbType.Int).Value = nTipoFteIngreso
        oComm.Parameters.Add("@cCodZona", SqlDbType.Char, 12).Value = sCodZona

        oComm.Parameters.Add("@bValidaParamFecha", SqlDbType.Bit).Value = bValidaParamFecha

        'oComm.ExecuteNonQuery()

        oComm.Parameters.Add("@nError", SqlDbType.Int)
        oComm.Parameters.Item("@nError").Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        nError = oComm.Parameters.Item("@nError").Value
        nCodEval = oComm.Parameters.Item("@nCodEval").Value

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvaluacion_Modificar(ByVal nCodEval As Integer, _
                        ByVal dFecEval As Date, _
                        ByVal nCodPers As Integer, _
                        ByVal sDescripcion As String, _
                        ByVal sNomNegocio As String, _
                        ByVal sDireccion As String, _
                        ByVal sTelefono As String, _
                        ByVal nTipoFteIngreso As Integer, _
                        ByVal sCodZona As String, _
                                        ByRef nResul As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvaluacion_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@dFecEval", SqlDbType.DateTime).Value = dFecEval
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 200).Value = sDescripcion
        oComm.Parameters.Add("@cNomNegocio", SqlDbType.VarChar, 50).Value = sNomNegocio
        oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 50).Value = sDireccion
        oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar, 20).Value = sTelefono
        oComm.Parameters.Add("@nTipoFteIngreso", SqlDbType.Int).Value = nTipoFteIngreso
        oComm.Parameters.Add("@cCodZona", SqlDbType.Char, 12).Value = sCodZona

        oComm.Parameters.Add("@nResul", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        nResul = oComm.Parameters("@nResul").Value

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Sub RegistraSupuestosFlujoCaja(ByVal nCodEval As Integer, ByVal psVentas As String, _
        ByVal psGastosOpe As String, ByVal psGastosFam As String, ByVal psOtros As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_RegistraSuspuestosFC", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCoEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@psVentas", SqlDbType.VarChar, 1000).Value = psVentas
        oComm.Parameters.Add("@psComGasOpe", SqlDbType.VarChar, 1000).Value = psGastosOpe
        oComm.Parameters.Add("@psComGasFam", SqlDbType.VarChar, 1000).Value = psGastosFam
        oComm.Parameters.Add("@psComOtros", SqlDbType.VarChar, 1000).Value = psOtros

        oComm.ExecuteNonQuery()


        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvaluacion_Eliminar(ByVal nCodEval As Integer, _
                                        ByRef nResul As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvaluacion_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nResul", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        nResul = oComm.Parameters("@nResul").Value

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperaBalanceGen(ByRef oDatos As ComTipos.clsDataset, _
             ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperaBalanceGen " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub PersEvalBalGen_Insertar(ByVal nCodEval As Integer, _
                            ByVal nDisponible As Double, _
                            ByVal nAdelantoProv As Double, _
                            ByVal nCtasCobrar As Double, _
                            ByVal nInventario As Double, _
                            ByVal nMercaderiaTran As Double, _
                            ByVal nBienesMueb As Double, _
                            ByVal nBienesInmue As Double, _
                            ByVal nCtasPagar As Double, _
                            ByVal nAdelantoCli As Double, _
                            ByVal nOtrasCtasPag As Double, _
                            ByVal nPasivoNC As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalBalGen_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = nDisponible
        oComm.Parameters.Add("@nAdelantoProv", SqlDbType.Money).Value = nAdelantoProv
        oComm.Parameters.Add("@nCtasCobrar", SqlDbType.Money).Value = nCtasCobrar
        oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = nInventario
        oComm.Parameters.Add("@nMercaderiaTran", SqlDbType.Money).Value = nMercaderiaTran
        oComm.Parameters.Add("@nBienesMueb", SqlDbType.Money).Value = nBienesMueb
        oComm.Parameters.Add("@nBienesInmue", SqlDbType.Money).Value = nBienesInmue
        oComm.Parameters.Add("@nCtasPagar", SqlDbType.Money).Value = nCtasPagar
        oComm.Parameters.Add("@nAdelantoCli", SqlDbType.Money).Value = nAdelantoCli
        oComm.Parameters.Add("@nOtrasCtasPag", SqlDbType.Money).Value = nOtrasCtasPag
        oComm.Parameters.Add("@nPasivoNoC", SqlDbType.Money).Value = nPasivoNC

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalBalGen_Modificar(ByVal nCodEval As Integer, _
                        ByVal nDisponible As Double, _
                        ByVal nAdelantoProv As Double, _
                        ByVal nCtasCobrar As Double, _
                        ByVal nInventario As Double, _
                        ByVal nMercaderiaTran As Double, _
                        ByVal nBienesMueb As Double, _
                        ByVal nBienesInmue As Double, _
                        ByVal nCtasPagar As Double, _
                        ByVal nAdelantoCli As Double, _
                        ByVal nOtrasCtasPag As Double, _
                        ByVal nPasivoNC As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalBalGen_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = nDisponible
        oComm.Parameters.Add("@nAdelantoProv", SqlDbType.Money).Value = nAdelantoProv
        oComm.Parameters.Add("@nCtasCobrar", SqlDbType.Money).Value = nCtasCobrar
        oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = nInventario
        oComm.Parameters.Add("@nMercaderiaTran", SqlDbType.Money).Value = nMercaderiaTran
        oComm.Parameters.Add("@nBienesMueb", SqlDbType.Money).Value = nBienesMueb
        oComm.Parameters.Add("@nBienesInmue", SqlDbType.Money).Value = nBienesInmue
        oComm.Parameters.Add("@nCtasPagar", SqlDbType.Money).Value = nCtasPagar
        oComm.Parameters.Add("@nAdelantoCli", SqlDbType.Money).Value = nAdelantoCli
        oComm.Parameters.Add("@nOtrasCtasPag", SqlDbType.Money).Value = nOtrasCtasPag
        oComm.Parameters.Add("@nPasivoNoC", SqlDbType.Money).Value = nPasivoNC

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    'Public Sub PersEvalBalGen_Grabar(ByVal nCodEval As Integer, _
    '                    ByVal nDisponible As Double, _
    '                    ByVal nAdelantoProv As Double, _
    '                    ByVal nCtasCobrar As Double, _
    '                    ByVal nInventario As Double, _
    '                    ByVal nMercaderiaTran As Double, _
    '                    ByVal nBienesMueb As Double, _
    '                    ByVal nBienesInmue As Double, _
    '                    ByVal nCtasPagar As Double, _
    '                    ByVal nAdelantoCli As Double, _
    '                    ByVal nOtrasCtasPag As Double, _
    '                    ByVal nPasivoNC As Double, _
    '                    ByVal nCtasPagarLP As Double, _
    '                    ByVal nOtrasCPagarLP As Double, _
    '                    ByVal nCapital As Double, _
    '                    ByVal nReservas As Double, _
    '                   ByVal nDistribUtilidad As Double, _
    '                   Optional ByVal nInversiones As Double = 0, _
    '                   Optional ByVal nDeudaActivoF As Double = 0)

    '    Dim objCon As ComConecta.clsConecta
    '    Dim oComm As SqlCommand

    '    objCon = New ComConecta.clsConecta
    '    oComm = New SqlCommand("Eval_PersEvalBalGen_Grabar", objCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
    '    oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = nDisponible
    '    oComm.Parameters.Add("@nAdelantoProv", SqlDbType.Money).Value = nAdelantoProv
    '    oComm.Parameters.Add("@nCtasCobrar", SqlDbType.Money).Value = nCtasCobrar
    '    oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = nInventario
    '    oComm.Parameters.Add("@nMercaderiaTran", SqlDbType.Money).Value = nMercaderiaTran
    '    oComm.Parameters.Add("@nBienesMueb", SqlDbType.Money).Value = nBienesMueb
    '    oComm.Parameters.Add("@nBienesInmue", SqlDbType.Money).Value = nBienesInmue
    '    oComm.Parameters.Add("@nCtasPagar", SqlDbType.Money).Value = nCtasPagar
    '    oComm.Parameters.Add("@nAdelantoCli", SqlDbType.Money).Value = nAdelantoCli
    '    oComm.Parameters.Add("@nOtrasCtasPag", SqlDbType.Money).Value = nOtrasCtasPag
    '    oComm.Parameters.Add("@nPasivoNoC", SqlDbType.Money).Value = nPasivoNC

    '    oComm.Parameters.Add("@nCtasPagarLP", SqlDbType.Money).Value = nCtasPagarLP
    '    oComm.Parameters.Add("@nOtrasCPagarLP", SqlDbType.Money).Value = nOtrasCPagarLP
    '    oComm.Parameters.Add("@nCapital", SqlDbType.Money).Value = nCapital
    '    oComm.Parameters.Add("@nReservas", SqlDbType.Money).Value = nReservas
    '    oComm.Parameters.Add("@nDistribUtilidad", SqlDbType.Money).Value = nDistribUtilidad

    '    oComm.Parameters.Add("@nInversiones", SqlDbType.Money).Value = nInversiones
    '    oComm.Parameters.Add("@nDeudaActivoF", SqlDbType.Money).Value = nDeudaActivoF
    '    oComm.ExecuteNonQuery()

    '    objCon = Nothing
    '    oComm = Nothing
    'End Sub
    Public Sub PersEvalBalGen_Grabar(ByVal nCodEval As Integer, _
                       ByVal nDisponible As Double, _
                       ByVal nAdelantoProv As Double, _
                       ByVal nCtasCobrar As Double, _
                       ByVal nInventario As Double, _
                       ByVal nMercaderiaTran As Double, _
                       ByVal nBienesMueb As Double, _
                       ByVal nBienesInmue As Double, _
                       ByVal nCtasPagar As Double, _
                       ByVal nAdelantoCli As Double, _
                       ByVal nOtrasCtasPag As Double, _
                       ByVal nPasivoNC As Double, _
                       ByVal nCtasPagarLP As Double, _
                       ByVal nOtrasCPagarLP As Double, _
                       ByVal nCapital As Double, _
                       ByVal nReservas As Double, _
                       ByVal nDistribUtilidad As Double, _
                       ByVal nMateriales As Double, _
                       ByVal nMatPrima As Double, _
                       ByVal nMatProceso As Double, _
                       ByVal nProdTerminado As Double, _
                       Optional ByVal nInversiones As Double = 0, _
                       Optional ByVal nDeudaActivoF As Double = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalBalGen_Grabar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = nDisponible
        oComm.Parameters.Add("@nAdelantoProv", SqlDbType.Money).Value = nAdelantoProv
        oComm.Parameters.Add("@nCtasCobrar", SqlDbType.Money).Value = nCtasCobrar
        oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = nInventario
        oComm.Parameters.Add("@nMercaderiaTran", SqlDbType.Money).Value = nMercaderiaTran
        oComm.Parameters.Add("@nBienesMueb", SqlDbType.Money).Value = nBienesMueb
        oComm.Parameters.Add("@nBienesInmue", SqlDbType.Money).Value = nBienesInmue
        oComm.Parameters.Add("@nCtasPagar", SqlDbType.Money).Value = nCtasPagar
        oComm.Parameters.Add("@nAdelantoCli", SqlDbType.Money).Value = nAdelantoCli
        oComm.Parameters.Add("@nOtrasCtasPag", SqlDbType.Money).Value = nOtrasCtasPag
        oComm.Parameters.Add("@nPasivoNoC", SqlDbType.Money).Value = nPasivoNC

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

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Sub PersEvalBalGen_Eliminar(ByVal nCodEval As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalBalGen_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperaEstadGananPerd(ByRef oDatos As ComTipos.clsDataset, _
         ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperaEstadGananPerd " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub PersEvalEstGanPerd_Insertar(ByVal nCodEval As Integer, _
                            ByVal nVentaContado As Double, _
                            ByVal nVentaCredito As Double, _
                            ByVal nPromDiasOtorCred As Integer, _
                            ByVal nCompraContado As Double, _
                            ByVal nCompraCredito As Double, _
                            ByVal nPromDiasPagCred As Integer, _
                            ByVal nVentasMensuales As Double, _
                            ByVal nCostoVentas As Double, _
                            ByVal nGastoPersonal As Double, _
                            ByVal nServicios As Double, _
                            ByVal nAlquilerLocal As Double, _
                            ByVal nTrasporteMer As Double, _
                            ByVal nImpuestos As Double, _
                            ByVal nGastosFinacNeg As Double, _
                            ByVal nOtrosGastos As Double, _
                            ByVal nOtrosIngresosFam As Double, _
                            ByVal nTotalGastosFam As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalEstGanPerd_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
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

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalEstGanPerd_Modificar(ByVal nCodEval As Integer, _
                         ByVal nVentaContado As Double, _
                         ByVal nVentaCredito As Double, _
                         ByVal nPromDiasOtorCred As Integer, _
                         ByVal nCompraContado As Double, _
                         ByVal nCompraCredito As Double, _
                         ByVal nPromDiasPagCred As Integer, _
                         ByVal nVentasMensuales As Double, _
                         ByVal nCostoVentas As Double, _
                         ByVal nGastoPersonal As Double, _
                         ByVal nServicios As Double, _
                         ByVal nAlquilerLocal As Double, _
                         ByVal nTrasporteMer As Double, _
                         ByVal nImpuestos As Double, _
                         ByVal nGastosFinacNeg As Double, _
                         ByVal nOtrosGastos As Double, _
                         ByVal nOtrosIngresosFam As Double, _
                         ByVal nTotalGastosFam As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalEstGanPerd_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
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

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalEstGanPerd_Grabar(ByVal nCodEval As Integer, _
                             ByVal nVentaContado As Double, _
                             ByVal nVentaCredito As Double, _
                             ByVal nPromDiasOtorCred As Integer, _
                             ByVal nCompraContado As Double, _
                             ByVal nCompraCredito As Double, _
                             ByVal nPromDiasPagCred As Integer, _
                             ByVal nVentasMensuales As Double, _
                             ByVal nCostoVentas As Double, _
                             ByVal nGastoPersonal As Double, _
                             ByVal nServicios As Double, _
                             ByVal nAlquilerLocal As Double, _
                             ByVal nTrasporteMer As Double, _
                             ByVal nImpuestos As Double, _
                             ByVal nGastosFinacNeg As Double, _
                             ByVal nOtrosGastos As Double, _
                             ByVal nOtrosIngresosFam As Double, _
                             ByVal nTotalGastosFam As Double, _
                             Optional ByVal nUtiNegoSecund As Double = 0, _
                             Optional ByVal nOtrosGastosFam As Double = 0, _
                             Optional ByVal bCostoManual As Boolean = False)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalEstGanPerd_Grabar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
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
        oComm.Parameters.Add("@nUtiNegSecundario", SqlDbType.Money).Value = nUtiNegoSecund
        oComm.Parameters.Add("@nOtrosGastosFam", SqlDbType.Money).Value = nOtrosGastosFam
        oComm.Parameters.Add("@bManual", SqlDbType.Bit).Value = bCostoManual

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalEstGanPerd_Eliminar(ByVal nCodEval As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalEstGanPerd_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperaIndicadoresFinancieros(ByRef oDatos As ComTipos.clsDataset, _
         ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperaIndicadoresFinancieros " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CargaIndicadoresFinancieros(ByRef oDatos As ComTipos.clsDataset, _
                        ByVal nCodEval As Integer, _
                        ByVal nPrestamo As Double, _
                        ByVal nCuota As Double, _
                        ByVal nParamRV As Double, _
                        ByVal nParamRI As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CargaIndicadoresFinancieros " & nCodEval & "," & nPrestamo & "," & nCuota & "," & nParamRV & "," & nParamRI)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub PersEvalIndFinanc_Insertar(ByVal nCodEval As Integer, _
                                    ByVal nPrestamo As Double, _
                                    ByVal nCuota As Double, _
                                    ByVal nParamRV As Double, _
                                    ByVal nParamRI As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalIndFinanc_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = nPrestamo
        oComm.Parameters.Add("@nCuota", SqlDbType.Money).Value = nCuota
        oComm.Parameters.Add("@nParamRV ", SqlDbType.Money).Value = nParamRV
        oComm.Parameters.Add("@nParamRI", SqlDbType.Money).Value = nParamRI

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalIndFinanc_Modificar(ByVal nCodEval As Integer, _
                                        ByVal nPrestamo As Double, _
                                        ByVal nCuota As Double, _
                                        ByVal nParamRV As Double, _
                                        ByVal nParamRI As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalIndFinanc_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = nPrestamo
        oComm.Parameters.Add("@nCuota", SqlDbType.Money).Value = nCuota
        oComm.Parameters.Add("@nParamRV ", SqlDbType.Money).Value = nParamRV
        oComm.Parameters.Add("@nParamRI", SqlDbType.Money).Value = nParamRI

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function ObtieneGastoTotalCal(ByVal pnPrestamo As Double, ByVal pnMoneda As Integer, _
        ByVal pnProd As Integer, ByVal pnSubProd As Integer, Optional ByVal pnSegInd As Integer = 1) As Double

        Dim R As ComTipos.clsDataset

        R = New ComTipos.clsDataset
        R.CargaTabla(, , "CredW_RecuperaGastosParaSimulador " & pnPrestamo & "," & pnMoneda & "," & pnProd & "," & pnSubProd & "," & pnSegInd)

        ObtieneGastoTotalCal = CDbl(R.ObtenerDato("nValorGasto"))

        R = Nothing

    End Function

    Public Function RecuperaFeriados(ByVal pdFechaDesem As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_DevuelveFeriados '" & Format(pdFechaDesem, "dd/MM/yyyy") & "'")
        RecuperaFeriados = D
        D = Nothing
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

    Public Sub PersEvalIndFinanc_Grabar(ByVal nCodEval As Integer, _
                                    ByVal nPrestamo As Double, _
                                    ByVal nCuota As Double, _
                                    ByVal nParamRV As Double, _
                                    ByVal nParamRI As Double, _
                                    ByVal nTEM As Double, ByVal pdFecDes As Date, ByVal pnNumcuotas As Integer, _
                                    ByVal pnTipoPer As Integer, ByVal pnPlazo As Integer, ByVal pnDiaVenc As Integer, _
                                    ByVal pnDiasGracia As Integer, ByVal pnTipoGracia As Integer, ByVal pnCodLin As Integer, _
                                    ByVal pnTasaCom As Double, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, _
                                    Optional ByVal nProducto As Integer = 0, _
                                    Optional ByVal nSubProducto As Integer = 0, _
                                    Optional ByVal pnItfValor As Double = 0, _
                                    Optional ByVal pnCodCamp As Integer = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset
        Dim nTasaDesgravamen As Double

        nMontoGasto = ObtieneGastoTotalCal(nPrestamo, 1, nProducto, nSubProducto)
        oFeriado = RecuperaFeriados(pdFecDes)

        'ARCV 21-05-2009
        nTasaDesgravamen = RecuperaTasaSeguroDesgravamen(pnCodCred, pnCodAge, _
                                    nPrestamo, 1, nProducto, nSubProducto)

        oCal = New ComTipos.ClsCalendario
        oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, pnCodAge, pnCodCred, _
                        1, nPrestamo, pnNumcuotas, pnTasaCom, _
                        pnDiaVenc, pdFecDes, pnDiasGracia, pnTipoGracia, 0, _
                        nProducto, nSubProducto, IIf(pnCodCamp = 0, True, False), pnTipoPer, pnPlazo, pnItfValor, , _
                         nTasaDesgravamen)

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalIndFinanc_Grabar", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("Eval_PersEvalIndFinanc_Grabar")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = nPrestamo
        oComm.Parameters.Add("@nCuota", SqlDbType.Money).Value = nCuota
        oComm.Parameters.Add("@nParamRV ", SqlDbType.Money).Value = nParamRV
        oComm.Parameters.Add("@nParamRI", SqlDbType.Money).Value = nParamRI
        oComm.Parameters.Add("@nTEM", SqlDbType.Money).Value = nTEM
        oComm.Parameters.Add("@pdFecDes", SqlDbType.DateTime).Value = pdFecDes
        oComm.Parameters.Add("@pnNumcuotas", SqlDbType.Money).Value = pnNumcuotas
        oComm.Parameters.Add("@pnTipoPer", SqlDbType.Int).Value = pnTipoPer
        oComm.Parameters.Add("@pnPlazo", SqlDbType.Int).Value = pnPlazo
        oComm.Parameters.Add("@pnDiaVenc", SqlDbType.Int).Value = pnDiaVenc
        oComm.Parameters.Add("@pnDiasGracia", SqlDbType.Int).Value = pnDiasGracia
        oComm.Parameters.Add("@pnTipoGracia", SqlDbType.Int).Value = pnTipoGracia
        oComm.Parameters.Add("@pnCodLin", SqlDbType.Int).Value = pnCodLin
        oComm.Parameters.Add("@pnTasaCom", SqlDbType.Money).Value = pnTasaCom
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred

        oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
        oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML

        'ARCV 11-05-2011
        oComm.ExecuteNonQuery()
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalIndFinanc_Eliminar(ByVal nCodEval As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalIndFinanc_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CargarInterfazVerifDomicilio(ByRef oPropiedad As ComTipos.clsDataset, _
                            ByRef oPropietario As ComTipos.clsDataset, _
                            ByRef oZona As ComTipos.clsDataset, _
                            ByRef oBarrio As ComTipos.clsDataset, _
                            ByRef oTipoViv As ComTipos.clsDataset, _
                            ByRef oMaterial As ComTipos.clsDataset, _
                            ByRef oEstadoViv As ComTipos.clsDataset)

        Dim C As ComConstante.clsConstante
        C = New ComConstante.clsConstante
        oPropiedad = C.RecuperaConstantes(2157)
        oPropietario = C.RecuperaConstantes(2158)
        oZona = C.RecuperaConstantes(2159)
        oBarrio = C.RecuperaConstantes(2160)
        oTipoViv = C.RecuperaConstantes(2161)
        oMaterial = C.RecuperaConstantes(2162)
        oEstadoViv = C.RecuperaConstantes(2163)
        C = Nothing
    End Sub

    Public Sub RecuperaVerificacionDomi(ByRef oDatos As ComTipos.clsPersEvalVerifDomic, _
                         ByVal nCodEval As Integer, _
                         Optional ByVal bAval As Boolean = False)
        'Eval_RecuperaVerificacionDomi

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        objCon = New ComConecta.clsConecta
        If bAval = False Then
            oComm = New SqlCommand("Eval_RecuperaVerificacionDomi", objCon.DameConexion)
        Else
            oComm = New SqlCommand("Eval_RecuperaVerificacionAval", objCon.DameConexion)
        End If
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        R = oComm.ExecuteReader
        'oComm.ExecuteNonQuery()

        While R.Read
            With oDatos
                .bVerificado = R("bVerificado")
                '.dFechaInstal = R("dFechaInstal")                
                .dFecReg = R("dFecReg")
                .nCodEval = R("nCodEval")
                .nCodPersVerif = R("nCodPersVerif")

                .nEstadoVivienda = R("nEstadoVivienda")
                .nTipoBarrio = R("nTipoBarrio")
                .nTipoMaterial = R("nTipoMaterial")
                .nTipoPropiedad = R("nTipoPropiedad")
                .nTipoPropietario = R("nTipoPropietario")
                .nTipoVivienda = R("nTipoVivienda")
                .nTipoZona = R("nTipoZona")
                .oFotoDomi = IIf(IsDBNull(R("oFotoDomi")), Nothing, R("oFotoDomi"))
                .oFotoNego = IIf(IsDBNull(R("oFotoNego")), Nothing, R("oFotoNego"))
                .sNomPersAten = R("cNomPersAten") 'ARCV 02-02-2010
                .sNomPersVerif = R("cNomPersVerif") 'ARCV 02-02-2010
                .sObservacion = R("cObservacion") 'ARCV 02-02-2010
                .sParentescoAten = R("cParentescoAten")
                .sReferenciaDomi = R("cReferenciaDomi")
                .sReferenciaNego = R("cReferenciaNego")
                .sTiempo = R("cTiempo")
                .oFotoDomi2 = IIf(IsDBNull(R("oFotoDomi2")), Nothing, R("oFotoDomi2"))
                .oFotoNego2 = IIf(IsDBNull(R("oFotoNego2")), Nothing, R("oFotoNego2"))
                .sReferenciaDomi2 = R("cReferenciaDomi2")
                .sReferenciaNego2 = R("cReferenciaNego2")

                If bAval Then
                    .sNomAval = R("cNomAval")
                    .sTelefAval = R("cTelefAval")
                    .sDirecAval = R("cDirecAval")
                End If

                'ARCV 23-11-2010
                .oFotoNego3 = IIf(IsDBNull(R("oFotoNego3")), Nothing, R("oFotoNego3"))
                .oFotoNego4 = IIf(IsDBNull(R("oFotoNego4")), Nothing, R("oFotoNego4"))
                .oFotoSumin = IIf(IsDBNull(R("oFotoSumin")), Nothing, R("oFotoSumin"))
                .sReferenciaNego3 = R("cReferenciaNego3")
                .sReferenciaNego4 = R("cReferenciaNego4")
                .sReferenciaSumin = R("cReferenciaSumin")
                '-------------
            End With
        End While
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperaFotosCloud(ByRef oDatos As ComTipos.clsPersEvalVerifDomic, ByVal nCodEval As Integer, ByVal nCodFoto As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        objCon = New ComConecta.clsConecta

        oComm = New SqlCommand("Pers_RecuperaImagen", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nCodFoto", SqlDbType.Int).Value = nCodFoto
        R = oComm.ExecuteReader
        'oComm.ExecuteNonQuery()

        While R.Read
            With oDatos
                Select Case nCodFoto
                    Case 1
                        .oFotoDomi = IIf(IsDBNull(R("oFotoDomi")), Nothing, R("oFotoDomi"))
                    Case 2
                        .oFotoDomi2 = IIf(IsDBNull(R("oFotoDomi2")), Nothing, R("oFotoDomi2"))
                    Case 3
                        .oFotoNego = IIf(IsDBNull(R("oFotoNego")), Nothing, R("oFotoNego"))
                    Case 4
                        .oFotoNego2 = IIf(IsDBNull(R("oFotoNego2")), Nothing, R("oFotoNego2"))
                    Case 5
                        .oFotoNego3 = IIf(IsDBNull(R("oFotoNego3")), Nothing, R("oFotoNego3"))
                    Case 6
                        .oFotoNego4 = IIf(IsDBNull(R("oFotoNego4")), Nothing, R("oFotoNego4"))
                    Case 7
                        .oFotoSumin = IIf(IsDBNull(R("oFotoSumin")), Nothing, R("oFotoSumin"))
                End Select
                '.oFotoDomi = IIf(IsBNull(R("oFotoDomi")), Nothing, R("oFotoDomi"))
                '.oFotoNego = IIf(IsDBNull(R("oFotoNego")), Nothing, R("oFotoNego"))
                '.sNomPersAten = R("cNomPersAten") 'ARCV 02-02-2010
                '.sNomPersVerif = R("cNomPersVerif") 'ARCV 02-02-2010
                '.sObservacion = R("cObservacion") 'ARCV 02-02-2010
                '.sParentescoAten = R("cParentescoAten")
                '.sReferenciaDomi = R("cReferenciaDomi")
                '.sReferenciaNego = R("cReferenciaNego")
                '.sTiempo = R("cTiempo")
                '.oFotoDomi2 = IIf(IsDBNull(R("oFotoDomi2")), Nothing, R("oFotoDomi2"))
                '.oFotoNego2 = IIf(IsDBNull(R("oFotoNego2")), Nothing, R("oFotoNego2"))
                '.sReferenciaDomi2 = R("cReferenciaDomi2")
                '.sReferenciaNego2 = R("cReferenciaNego2")

                'If bAval Then
                '    .sNomAval = R("cNomAval")
                '    .sTelefAval = R("cTelefAval")
                '    .sDirecAval = R("cDirecAval")
                'End If

                ''ARCV 23-11-2010
                '.oFotoNego3 = IIf(IsDBNull(R("oFotoNego3")), Nothing, R("oFotoNego3"))
                '.oFotoNego4 = IIf(IsDBNull(R("oFotoNego4")), Nothing, R("oFotoNego4"))
                '.oFotoSumin = IIf(IsDBNull(R("oFotoSumin")), Nothing, R("oFotoSumin"))
                '.sReferenciaNego3 = R("cReferenciaNego3")
                '.sReferenciaNego4 = R("cReferenciaNego4")
                '.sReferenciaSumin = R("cReferenciaSumin")
                '-------------
            End With
        End While
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalVerifDomic_Insertar(ByVal nCodEval As Integer, _
                             ByVal dFecReg As Date, _
                             ByVal oFotoDomi() As Byte, _
                             ByVal sReferenciaDomi As String, _
                             ByVal oFotoNego() As Byte, _
                             ByVal sReferenciaNego As String, _
                             ByVal bVerificado As Integer, _
                             ByVal sTiempo As String, _
                             ByVal dFechaInstal As Date, _
                             ByVal nTipoPropiedad As Integer, _
                             ByVal nTipoPropietario As Integer, _
                             ByVal nTipoZona As Integer, _
                             ByVal nTipoBarrio As Integer, _
                             ByVal nTipoVivienda As Integer, _
                             ByVal nTipoMaterial As Integer, _
                             ByVal nEstadoVivienda As Integer, _
                             ByVal sObservacion As String, _
                             ByVal nCodPersVerif As Integer, _
                             ByVal sNomPersAten As String, _
                             ByVal sParentescoAten As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalVerifDomic_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@oFotoDomi", SqlDbType.Image).Value = oFotoDomi
        oComm.Parameters.Add("@cReferenciaDomi", SqlDbType.VarChar, 200).Value = sReferenciaDomi
        oComm.Parameters.Add("@oFotoNego", SqlDbType.Image).Value = oFotoNego
        oComm.Parameters.Add("@cReferenciaNego", SqlDbType.VarChar, 200).Value = sReferenciaNego
        oComm.Parameters.Add("@bVerificado", SqlDbType.Bit).Value = bVerificado
        oComm.Parameters.Add("@cTiempo", SqlDbType.VarChar, 20).Value = sTiempo
        oComm.Parameters.Add("@dFechaInstal", SqlDbType.DateTime).Value = dFechaInstal
        oComm.Parameters.Add("@nTipoPropiedad", SqlDbType.Int).Value = nTipoPropiedad
        oComm.Parameters.Add("@nTipoPropietario", SqlDbType.Int).Value = nTipoPropietario
        oComm.Parameters.Add("@nTipoZona", SqlDbType.Int).Value = nTipoZona
        oComm.Parameters.Add("@nTipoBarrio", SqlDbType.Int).Value = nTipoBarrio
        oComm.Parameters.Add("@nTipoVivienda", SqlDbType.Int).Value = nTipoVivienda
        oComm.Parameters.Add("@nTipoMaterial", SqlDbType.Int).Value = nTipoMaterial
        oComm.Parameters.Add("@nEstadoVivienda", SqlDbType.Int).Value = nEstadoVivienda
        oComm.Parameters.Add("@cObservacion", SqlDbType.VarChar, 400).Value = sObservacion
        oComm.Parameters.Add("@nCodPersVerif", SqlDbType.Int).Value = nCodPersVerif
        oComm.Parameters.Add("@cNomPersAten", SqlDbType.VarChar, 50).Value = sNomPersAten
        oComm.Parameters.Add("@cParentescoAten", SqlDbType.VarChar, 20).Value = sParentescoAten

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalVerifDomic_Modificar(ByVal nCodEval As Integer, _
                          ByVal dFecReg As Date, _
                          ByVal oFotoDomi() As Byte, _
                          ByVal sReferenciaDomi As String, _
                          ByVal oFotoNego() As Byte, _
                          ByVal sReferenciaNego As String, _
                          ByVal bVerificado As Integer, _
                          ByVal sTiempo As String, _
                          ByVal dFechaInstal As Date, _
                          ByVal nTipoPropiedad As Integer, _
                          ByVal nTipoPropietario As Integer, _
                          ByVal nTipoZona As Integer, _
                          ByVal nTipoBarrio As Integer, _
                          ByVal nTipoVivienda As Integer, _
                          ByVal nTipoMaterial As Integer, _
                          ByVal nEstadoVivienda As Integer, _
                          ByVal sObservacion As String, _
                          ByVal nCodPersVerif As Integer, _
                          ByVal sNomPersAten As String, _
                          ByVal sParentescoAten As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalVerifDomic_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@oFotoDomi", SqlDbType.Image).Value = oFotoDomi
        oComm.Parameters.Add("@cReferenciaDomi", SqlDbType.VarChar, 200).Value = sReferenciaDomi
        oComm.Parameters.Add("@oFotoNego", SqlDbType.Image).Value = oFotoNego
        oComm.Parameters.Add("@cReferenciaNego", SqlDbType.VarChar, 200).Value = sReferenciaNego
        oComm.Parameters.Add("@bVerificado", SqlDbType.Bit).Value = bVerificado
        oComm.Parameters.Add("@cTiempo", SqlDbType.VarChar, 20).Value = sTiempo
        oComm.Parameters.Add("@dFechaInstal", SqlDbType.DateTime).Value = dFechaInstal
        oComm.Parameters.Add("@nTipoPropiedad", SqlDbType.Int).Value = nTipoPropiedad
        oComm.Parameters.Add("@nTipoPropietario", SqlDbType.Int).Value = nTipoPropietario
        oComm.Parameters.Add("@nTipoZona", SqlDbType.Int).Value = nTipoZona
        oComm.Parameters.Add("@nTipoBarrio", SqlDbType.Int).Value = nTipoBarrio
        oComm.Parameters.Add("@nTipoVivienda", SqlDbType.Int).Value = nTipoVivienda
        oComm.Parameters.Add("@nTipoMaterial", SqlDbType.Int).Value = nTipoMaterial
        oComm.Parameters.Add("@nEstadoVivienda", SqlDbType.Int).Value = nEstadoVivienda
        oComm.Parameters.Add("@cObservacion", SqlDbType.VarChar, 400).Value = sObservacion
        oComm.Parameters.Add("@nCodPersVerif", SqlDbType.Int).Value = nCodPersVerif
        oComm.Parameters.Add("@cNomPersAten", SqlDbType.VarChar, 50).Value = sNomPersAten
        oComm.Parameters.Add("@cParentescoAten", SqlDbType.VarChar, 20).Value = sParentescoAten

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalVerifDomic_GrabarCloud(ByVal nCodEval As Integer, _
                           ByVal sReferenciaDomi As String, _
                           ByVal sReferenciaNego As String, _
                           ByVal sReferenciaDomi2 As String, _
                           ByVal sReferenciaNego2 As String, _
                       ByVal sReferenciaNego3 As String, _
                       ByVal sReferenciaNego4 As String, _
                       ByVal sReferenciaSumin As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalDomic_Referencias", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@cReferenciaDomi", SqlDbType.VarChar, 200).Value = sReferenciaDomi
        oComm.Parameters.Add("@cReferenciaNego", SqlDbType.VarChar, 200).Value = sReferenciaNego
        oComm.Parameters.Add("@cReferenciaDomi2", SqlDbType.VarChar, 200).Value = sReferenciaDomi2
        oComm.Parameters.Add("@cReferenciaNego2", SqlDbType.VarChar, 200).Value = sReferenciaNego2
        oComm.Parameters.Add("@cReferenciaNego3", SqlDbType.VarChar, 200).Value = sReferenciaNego3
        oComm.Parameters.Add("@cReferenciaNego4", SqlDbType.VarChar, 200).Value = sReferenciaNego4
        oComm.Parameters.Add("@cReferenciaSumin", SqlDbType.VarChar, 200).Value = sReferenciaSumin
        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing

    End Sub


    Public Sub PersEvalVerifDomic_Grabar(ByVal nCodEval As Integer, _
                             ByVal dFecReg As Date, _
                             ByVal oFotoDomi() As Byte, _
                             ByVal sReferenciaDomi As String, _
                             ByVal oFotoNego() As Byte, _
                             ByVal sReferenciaNego As String, _
                             ByVal bVerificado As Integer, _
                             ByVal sTiempo As String, _
                             ByVal dFechaInstal As Date, _
                             ByVal nTipoPropiedad As Integer, _
                             ByVal nTipoPropietario As Integer, _
                             ByVal nTipoZona As Integer, _
                             ByVal nTipoBarrio As Integer, _
                             ByVal nTipoVivienda As Integer, _
                             ByVal nTipoMaterial As Integer, _
                             ByVal nEstadoVivienda As Integer, _
                             ByVal sObservacion As String, _
                             ByVal nCodPersVerif As Integer, _
                             ByVal sNomPersAten As String, _
                             ByVal sParentescoAten As String, _
                             ByVal oFotoDomi2() As Byte, _
                             ByVal sReferenciaDomi2 As String, _
                             ByVal oFotoNego2() As Byte, _
                             ByVal sReferenciaNego2 As String, _
                         ByVal oFotoNego3() As Byte, _
                         ByVal sReferenciaNego3 As String, _
                         ByVal oFotoNego4() As Byte, _
                         ByVal sReferenciaNego4 As String, _
                         ByVal oFotoSumin() As Byte, _
                         ByVal sReferenciaSumin As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Eval_PersEvalVerifDomic_Grabar", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Eval_PersEvalVerifDomic_Grabar")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@oFotoDomi", SqlDbType.Image).Value = IIf(oFotoDomi Is Nothing, DBNull.Value, oFotoDomi)
        oComm.Parameters.Add("@cReferenciaDomi", SqlDbType.VarChar, 200).Value = sReferenciaDomi
        oComm.Parameters.Add("@oFotoNego", SqlDbType.Image).Value = IIf(oFotoNego Is Nothing, DBNull.Value, oFotoNego)
        oComm.Parameters.Add("@cReferenciaNego", SqlDbType.VarChar, 200).Value = sReferenciaNego
        oComm.Parameters.Add("@bVerificado", SqlDbType.Bit).Value = bVerificado
        oComm.Parameters.Add("@cTiempo", SqlDbType.VarChar, 20).Value = sTiempo
        oComm.Parameters.Add("@dFechaInstal", SqlDbType.DateTime).Value = dFechaInstal
        oComm.Parameters.Add("@nTipoPropiedad", SqlDbType.Int).Value = nTipoPropiedad
        oComm.Parameters.Add("@nTipoPropietario", SqlDbType.Int).Value = nTipoPropietario
        oComm.Parameters.Add("@nTipoZona", SqlDbType.Int).Value = nTipoZona
        oComm.Parameters.Add("@nTipoBarrio", SqlDbType.Int).Value = nTipoBarrio
        oComm.Parameters.Add("@nTipoVivienda", SqlDbType.Int).Value = nTipoVivienda
        oComm.Parameters.Add("@nTipoMaterial", SqlDbType.Int).Value = nTipoMaterial
        oComm.Parameters.Add("@nEstadoVivienda", SqlDbType.Int).Value = nEstadoVivienda
        oComm.Parameters.Add("@cObservacion", SqlDbType.VarChar, 400).Value = sObservacion
        oComm.Parameters.Add("@nCodPersVerif", SqlDbType.Int).Value = nCodPersVerif
        oComm.Parameters.Add("@cNomPersAten", SqlDbType.VarChar, 50).Value = sNomPersAten
        oComm.Parameters.Add("@cParentescoAten", SqlDbType.VarChar, 20).Value = sParentescoAten
        oComm.Parameters.Add("@oFotoDomi2", SqlDbType.Image).Value = IIf(oFotoDomi2 Is Nothing, DBNull.Value, oFotoDomi2)
        oComm.Parameters.Add("@cReferenciaDomi2", SqlDbType.VarChar, 200).Value = sReferenciaDomi2
        oComm.Parameters.Add("@oFotoNego2", SqlDbType.Image).Value = IIf(oFotoNego2 Is Nothing, DBNull.Value, oFotoNego2)
        oComm.Parameters.Add("@cReferenciaNego2", SqlDbType.VarChar, 200).Value = sReferenciaNego2

        'ARCV 23-11-2010
        oComm.Parameters.Add("@oFotoNego3", SqlDbType.Image).Value = IIf(oFotoNego3 Is Nothing, DBNull.Value, oFotoNego3)
        oComm.Parameters.Add("@cReferenciaNego3", SqlDbType.VarChar, 200).Value = sReferenciaNego3
        oComm.Parameters.Add("@oFotoNego4", SqlDbType.Image).Value = IIf(oFotoNego4 Is Nothing, DBNull.Value, oFotoNego4)
        oComm.Parameters.Add("@cReferenciaNego4", SqlDbType.VarChar, 200).Value = sReferenciaNego4
        oComm.Parameters.Add("@oFotoSumin", SqlDbType.Image).Value = IIf(oFotoSumin Is Nothing, DBNull.Value, oFotoSumin)
        oComm.Parameters.Add("@cReferenciaSumin", SqlDbType.VarChar, 200).Value = sReferenciaSumin
        '---------------

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalVerifDomic_Eliminar(ByVal nCodEval As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalVerifDomic_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CargarEvaluacion_Impresion(ByRef oDatos As ComTipos.clsDataset, _
        ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CargarEvaluacion_Impresion " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Function RecuperaCIIU(ByVal nCodPers As Integer) As String

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "PersEval_REcuperaCIIU " & nCodPers)
        Return oData.ObtenerDato("cNomCIIU")
        oData = Nothing
    End Function

    Public Sub CalculaEstadGananPerd(ByRef oDatos As ComTipos.clsDataset, _
    ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CalculaEstadGananPerd " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CalculaBalanceGen(ByRef oDatos As ComTipos.clsDataset, _
        ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CalculaBalanceGen " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Function RecuperaAnalistas(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CredW_RecuperaAnalista  " & Trim(Str(nCodAge)))
        Return oData
    End Function

    Public Function RecuperaDatosBasicos_Credito(ByVal nCodAge As Integer, _
                                                ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaDatosBasicos_Credito  " & nCodAge & "," & nCodCred)
        Return oData
    End Function

    Public Function CargaEvaluacion_x_Expediente(ByVal nCodAge As Integer, _
                                            ByVal nCodCred As Integer) As Integer
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CargaEvaluacion_x_Expediente  " & nCodAge & "," & nCodCred)
        If oData.nNumFilas = 0 Then
            Return 0
        Else
            Return CDbl(oData.ObtenerDato("nCodEval"))
        End If
    End Function

    Public Sub PersEvalVerifAval_Grabar(ByVal nCodEval As Integer, _
                         ByVal dFecReg As Date, _
                         ByVal oFotoDomi() As Byte, _
                         ByVal sReferenciaDomi As String, _
                         ByVal oFotoNego() As Byte, _
                         ByVal sReferenciaNego As String, _
                         ByVal bVerificado As Integer, _
                         ByVal sTiempo As String, _
                         ByVal dFechaInstal As Date, _
                         ByVal nTipoPropiedad As Integer, _
                         ByVal nTipoPropietario As Integer, _
                         ByVal nTipoZona As Integer, _
                         ByVal nTipoBarrio As Integer, _
                         ByVal nTipoVivienda As Integer, _
                         ByVal nTipoMaterial As Integer, _
                         ByVal nEstadoVivienda As Integer, _
                         ByVal sObservacion As String, _
                         ByVal nCodPersVerif As Integer, _
                         ByVal sNomPersAten As String, _
                         ByVal sParentescoAten As String, _
                         ByVal nCodPersAval As Integer, _
                         ByVal oFotoDomi2() As Byte, _
                         ByVal sReferenciaDomi2 As String, _
                         ByVal oFotoNego2() As Byte, _
                         ByVal sReferenciaNego2 As String, _
                         ByVal oFotoNego3() As Byte, _
                         ByVal sReferenciaNego3 As String, _
                         ByVal oFotoNego4() As Byte, _
                         ByVal sReferenciaNego4 As String, _
                         ByVal oFotoSumin() As Byte, _
                         ByVal sReferenciaSumin As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Eval_PersEvalVerifAval_Grabar", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Eval_PersEvalVerifAval_Grabar")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = dFecReg
        oComm.Parameters.Add("@oFotoDomi", SqlDbType.Image).Value = IIf(oFotoDomi Is Nothing, DBNull.Value, oFotoDomi) ' oFotoDomi
        oComm.Parameters.Add("@cReferenciaDomi", SqlDbType.VarChar, 200).Value = sReferenciaDomi
        oComm.Parameters.Add("@oFotoNego", SqlDbType.Image).Value = IIf(oFotoNego Is Nothing, DBNull.Value, oFotoNego)
        oComm.Parameters.Add("@cReferenciaNego", SqlDbType.VarChar, 200).Value = sReferenciaNego
        oComm.Parameters.Add("@bVerificado", SqlDbType.Bit).Value = bVerificado
        oComm.Parameters.Add("@cTiempo", SqlDbType.VarChar, 20).Value = sTiempo
        oComm.Parameters.Add("@dFechaInstal", SqlDbType.DateTime).Value = dFechaInstal
        oComm.Parameters.Add("@nTipoPropiedad", SqlDbType.Int).Value = nTipoPropiedad
        oComm.Parameters.Add("@nTipoPropietario", SqlDbType.Int).Value = nTipoPropietario
        oComm.Parameters.Add("@nTipoZona", SqlDbType.Int).Value = nTipoZona
        oComm.Parameters.Add("@nTipoBarrio", SqlDbType.Int).Value = nTipoBarrio
        oComm.Parameters.Add("@nTipoVivienda", SqlDbType.Int).Value = nTipoVivienda
        oComm.Parameters.Add("@nTipoMaterial", SqlDbType.Int).Value = nTipoMaterial
        oComm.Parameters.Add("@nEstadoVivienda", SqlDbType.Int).Value = nEstadoVivienda
        oComm.Parameters.Add("@cObservacion", SqlDbType.VarChar, 400).Value = sObservacion
        oComm.Parameters.Add("@nCodPersVerif", SqlDbType.Int).Value = nCodPersVerif
        oComm.Parameters.Add("@cNomPersAten", SqlDbType.VarChar, 50).Value = sNomPersAten
        oComm.Parameters.Add("@cParentescoAten", SqlDbType.VarChar, 20).Value = sParentescoAten
        oComm.Parameters.Add("@nCodPersAval", SqlDbType.Int).Value = nCodPersAval
        oComm.Parameters.Add("@oFotoDomi2", SqlDbType.Image).Value = IIf(oFotoDomi2 Is Nothing, DBNull.Value, oFotoDomi2)
        oComm.Parameters.Add("@cReferenciaDomi2", SqlDbType.VarChar, 200).Value = sReferenciaDomi2
        oComm.Parameters.Add("@oFotoNego2", SqlDbType.Image).Value = IIf(oFotoNego2 Is Nothing, DBNull.Value, oFotoNego2)
        oComm.Parameters.Add("@cReferenciaNego2", SqlDbType.VarChar, 200).Value = sReferenciaNego2

        'ARCV 23-11-2010
        oComm.Parameters.Add("@oFotoNego3", SqlDbType.Image).Value = IIf(oFotoNego3 Is Nothing, DBNull.Value, oFotoNego3)
        oComm.Parameters.Add("@cReferenciaNego3", SqlDbType.VarChar, 200).Value = sReferenciaNego3
        oComm.Parameters.Add("@oFotoNego4", SqlDbType.Image).Value = IIf(oFotoNego4 Is Nothing, DBNull.Value, oFotoNego4)
        oComm.Parameters.Add("@cReferenciaNego4", SqlDbType.VarChar, 200).Value = sReferenciaNego4
        oComm.Parameters.Add("@oFotoSumin", SqlDbType.Image).Value = IIf(oFotoSumin Is Nothing, DBNull.Value, oFotoSumin)
        oComm.Parameters.Add("@cReferenciaSumin", SqlDbType.VarChar, 200).Value = sReferenciaSumin
        '---------------

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function BuscaExpediente_x_CodCliente(ByVal nCodPers As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_BuscaExpediente_x_CodCliente  " & nCodPers)
        Return oData

    End Function

    Public Sub RecuperaGestorCobranzasNegociacionesFIDEICOMISO(ByRef oDatos As ComTipos.clsDataset, _
            ByVal nCodAge As Integer, ByVal nCodCred As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "DBCORE..Prejudicial_RecuperaGestorCobranzasNegociaciones " & nCodAge & "," & nCodCred)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CargaIndicadoresCampana(ByRef oDatos As ComTipos.clsDataset, _
                    ByVal nCodEval As Integer, _
                    ByVal nPrestamo As Double, _
                    ByVal nUtilidad As Double, _
                    ByVal nIntereses As Double, _
                    ByVal nVentas As Double, _
                    ByVal nInventarios As Double)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CargaIndicadoresCampana " & nCodEval & "," & nPrestamo & "," & nUtilidad & "," & nIntereses & _
                        "," & nVentas & "," & nInventarios)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub PersEvalIndCampana_Grabar(ByVal nCodEval As Integer, _
                                ByVal nVentas As Double, _
                                ByVal nCostoVentas As Double, _
                                ByVal nCostosOperativos As Double, _
                                ByVal nInventarios As Double, _
                                ByVal cComentNegocio As String, _
                                ByVal cComentComportCred As String, _
                                ByVal cComentGarantias As String, _
                                ByVal cComentPropuesta As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalIndCampana_Grabar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nVentas", SqlDbType.Money).Value = nVentas
        oComm.Parameters.Add("@nCostoVentas", SqlDbType.Money).Value = nCostoVentas
        oComm.Parameters.Add("@nCostosOperativos ", SqlDbType.Money).Value = nCostosOperativos
        oComm.Parameters.Add("@nInventarios", SqlDbType.Money).Value = nInventarios
        oComm.Parameters.Add("@cComentNegocio", SqlDbType.VarChar, 250).Value = cComentNegocio
        oComm.Parameters.Add("@cComentComportCred", SqlDbType.VarChar, 250).Value = cComentComportCred
        oComm.Parameters.Add("@cComentGarantias", SqlDbType.VarChar, 250).Value = cComentGarantias
        oComm.Parameters.Add("@cComentPropuesta", SqlDbType.VarChar, 250).Value = cComentPropuesta


        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalModalidad_Grabar(ByVal nCodEval As Integer, _
                                ByVal nVentas As Double, _
                                ByVal nCostoVentas As Double, _
                                ByVal nCostosOperativos As Double, _
                                ByVal nPromCred As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalIndModalidad_Grabar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nVentas", SqlDbType.Money).Value = nVentas
        oComm.Parameters.Add("@nCostoVentas", SqlDbType.Money).Value = nCostoVentas
        oComm.Parameters.Add("@nCostosOperativos ", SqlDbType.Money).Value = nCostosOperativos
        oComm.Parameters.Add("@nPromCred", SqlDbType.Money).Value = nPromCred


        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperaIndicadoresCampana(ByRef oDatos As ComTipos.clsDataset, _
     ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperaIndicadoresCampana " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub RecuperaIndicadoresModalidad(ByRef oDatos As ComTipos.clsDataset, _
     ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperaIndicadoresModalidad " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub
    Public Function CargaDatosEvaluacionGrupal(ByVal pnCodGrupo As Integer, ByVal pnCiclo As Integer) As ComTipos.ClsConsGrupal
        Dim D As ComTipos.ClsConsGrupal
        D = New ComTipos.ClsConsGrupal
        D.oIntegrantes.CargaTabla(, , "Cred_RecuperaEvaluacionGrpIntegrante " & pnCodGrupo & "," & pnCiclo)
        If D.oIntegrantes.nNumFilas > 0 Then
            D.oData.CargaTabla(, , "Cred_RecuperaDatosConsGrp " & D.oIntegrantes.ObtenerDato("nCodCred") & "," & D.oIntegrantes.ObtenerDato("nCodAge"))
        End If

        CargaDatosEvaluacionGrupal = D
    End Function

    Public Sub PersEvalActualizaEvaluacionGrupal(ByVal nCodEval As Integer,
                            ByVal nCodGrupo As Integer,
                            ByVal nCiclo As Integer,
                            ByVal dFecha As Date,
                            ByVal oIntegran As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Cred_Graba_GarantiaGrupal", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Eval_PersEvalActualizaEvaluacionGrupal")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nCodGrupo", SqlDbType.Int).Value = nCodGrupo
        oComm.Parameters.Add("@nCiclo", SqlDbType.Int).Value = nCiclo
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@oIntegran ", SqlDbType.Xml).Value = oIntegran.DevuelveXML

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub


    Public Sub App_GrabaCreditoDiario(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer,
        nTotalventas As Double, nTotalcompras As Double, nGanancia As Double, nInventario As Double, nMontosol As Double,
        nPlazo As Double, nCuota As Double, nTotalventaslun As Double,
        nTotalcompraslun As Double, ncajalun As Double, nTotalventasmar As Double, nTotalcomprasmar As Double, ncajamar As Double, nTotalventasmie As Double,
        nTotalcomprasmie As Double, ncajamie As Double, nTotalventasjue As Double, nTotalcomprasjue As Double, ncajajue As Double, nTotalventasvie As Double,
        nTotalcomprasvie As Double, ncajavie As Double, nTotalventassab As Double, nTotalcomprassab As Double,
        ncajasab As Double, nTotalventasdom As Double, nTotalcomprasdom As Double, ncajadom As Double, nalimentacion As Double, nalquiler As Double,
        nvestido As Double, neducacion As Double, nsalud As Double, ntransporte As Double, nservicios As Double, npension As Double, nrecreacion As Double,
        notros As Double, ncuotatarjeta As Double, ncuotacasacom As Double, ncuotapresifi As Double, ncuotaprestamista As Double,
        ncuotajuntas As Double, notrosingr As Double, notrosegresos As Double)


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("App_GrabaCreditoDiario")
        oComm.Connection = objCon.DameConexion
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nTotalventas", SqlDbType.Money).Value = nTotalventas
        oComm.Parameters.Add("@nTotalcompras", SqlDbType.Money).Value = nTotalcompras
        oComm.Parameters.Add("@nGanancia", SqlDbType.Money).Value = nGanancia
        oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = nInventario
        oComm.Parameters.Add("@nMontosol", SqlDbType.Money).Value = nMontosol
        oComm.Parameters.Add("@nPlazo", SqlDbType.Money).Value = nPlazo
        oComm.Parameters.Add("@nCuota", SqlDbType.Money).Value = nCuota
        oComm.Parameters.Add("@nTotalventaslun", SqlDbType.Money).Value = nTotalventaslun
        oComm.Parameters.Add("@nTotalcompraslun", SqlDbType.Money).Value = nTotalcompraslun
        oComm.Parameters.Add("@ncajalun", SqlDbType.Money).Value = ncajalun
        oComm.Parameters.Add("@nTotalventasmar", SqlDbType.Money).Value = nTotalventasmar
        oComm.Parameters.Add("@nTotalcomprasmar", SqlDbType.Money).Value = nTotalcomprasmar
        oComm.Parameters.Add("@ncajamar", SqlDbType.Money).Value = ncajamar
        oComm.Parameters.Add("@nTotalventasmie", SqlDbType.Money).Value = nTotalventasmie
        oComm.Parameters.Add("@nTotalcomprasmie", SqlDbType.Money).Value = nTotalcomprasmie
        oComm.Parameters.Add("@ncajamie", SqlDbType.Money).Value = ncajamie
        oComm.Parameters.Add("@nTotalventasjue", SqlDbType.Money).Value = nTotalventasjue
        oComm.Parameters.Add("@nTotalcomprasjue", SqlDbType.Money).Value = nTotalcomprasjue
        oComm.Parameters.Add("@ncajajue", SqlDbType.Money).Value = ncajajue
        oComm.Parameters.Add("@nTotalventasvie", SqlDbType.Money).Value = nTotalventasvie
        oComm.Parameters.Add("@nTotalcomprasvie", SqlDbType.Money).Value = nTotalcomprasvie
        oComm.Parameters.Add("@ncajavie", SqlDbType.Money).Value = ncajavie
        oComm.Parameters.Add("@nTotalventassab", SqlDbType.Money).Value = nTotalventassab
        oComm.Parameters.Add("@nTotalcomprassab", SqlDbType.Money).Value = nTotalcomprassab
        oComm.Parameters.Add("@ncajasab", SqlDbType.Money).Value = ncajasab
        oComm.Parameters.Add("@nTotalventasdom", SqlDbType.Money).Value = nTotalventasdom
        oComm.Parameters.Add("@nTotalcomprasdom", SqlDbType.Money).Value = nTotalcomprasdom
        oComm.Parameters.Add("@ncajadom", SqlDbType.Money).Value = ncajadom
        oComm.Parameters.Add("@nalimentacion", SqlDbType.Money).Value = nalimentacion
        oComm.Parameters.Add("@nalquiler", SqlDbType.Money).Value = nalquiler
        oComm.Parameters.Add("@nvestido", SqlDbType.Money).Value = nvestido
        oComm.Parameters.Add("@neducacion", SqlDbType.Money).Value = neducacion
        oComm.Parameters.Add("@nsalud", SqlDbType.Money).Value = nsalud
        oComm.Parameters.Add("@ntransporte", SqlDbType.Money).Value = ntransporte
        oComm.Parameters.Add("@nservicios", SqlDbType.Money).Value = nservicios
        oComm.Parameters.Add("@npension", SqlDbType.Money).Value = npension
        oComm.Parameters.Add("@nrecreacion", SqlDbType.Money).Value = nrecreacion
        oComm.Parameters.Add("@notros", SqlDbType.Money).Value = notros
        oComm.Parameters.Add("@ncuotatarjeta", SqlDbType.Money).Value = ncuotatarjeta
        oComm.Parameters.Add("@ncuotacasacom", SqlDbType.Money).Value = ncuotacasacom
        oComm.Parameters.Add("@ncuotapresifi", SqlDbType.Money).Value = ncuotapresifi
        oComm.Parameters.Add("@ncuotaprestamista", SqlDbType.Money).Value = ncuotaprestamista
        oComm.Parameters.Add("@ncuotajuntas", SqlDbType.Money).Value = ncuotajuntas
        oComm.Parameters.Add("@notrosingr", SqlDbType.Money).Value = notrosingr
        oComm.Parameters.Add("@notrosegresos", SqlDbType.Money).Value = notrosegresos

        oComm.ExecuteNonQuery()


        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Function RecuperaEvalCredDiarioNuevo(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsWRegistroCredDiario
        Dim D As ComTipos.ClsWRegistroCredDiario
        D = New ComTipos.ClsWRegistroCredDiario
        D.Datos.CargaTabla(, , "App_RecuperaEvaluacionDiaria " & pnCodAge & "," & pnCodCred)
        RecuperaEvalCredDiarioNuevo = D
        D = Nothing
    End Function


    Public Function RecuperaEvalCredDiario(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsWRegistroCredDiario
        Dim D As ComTipos.ClsWRegistroCredDiario
        D = New ComTipos.ClsWRegistroCredDiario
        D.Datos.CargaTabla(, , "Cred_RecuperaDatosCalendarioEvalCredDiario " & pnCodAge & "," & pnCodCred)
        If D.Datos.nNumFilas > 0 Then
            D.CboClasificacion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2166")
            D.Deudas.CargaTabla(, , "Cred_RecuperaDatosEvalDiario_Deudas " & D.Datos.ObtenerDato("nCodEval") & "," & D.Datos.ObtenerDato("cDni"))
        End If
        RecuperaEvalCredDiario = D
        D = Nothing
    End Function

    Public Function RecuperaEvalConsumo(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.ClsWRegistroCredConsumo
        Dim D As ComTipos.ClsWRegistroCredConsumo
        D = New ComTipos.ClsWRegistroCredConsumo

        D.Datos.CargaTabla(, , "Cred_RecuperaDatosCalendarioEvalConsumo " & pnCodAge & "," & pnCodCred)
        D.CboCalificacion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 4550")

        D.cboTipoGrati.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 95") 'ARCV 10-09-2010

        If D.Datos.nNumFilas > 0 Then
            D.CboClasificacion.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2166")
            D.Deudas.CargaTabla(, , "Cred_RecuperaDatosEvalConsumo_Deudas " & D.Datos.ObtenerDato("nCodEval") & "," & D.Datos.ObtenerDato("cDni"))
        End If
        RecuperaEvalConsumo = D
        D = Nothing

    End Function

    Public Function RecuperaEvalRangosConsumo(ByVal nTotal As Double, ByVal nPorc As Double) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevueleRangosEvalConsumo " & nTotal & ", " & nPorc)
        RecuperaEvalRangosConsumo = D
        D = Nothing
    End Function
    Public Function RecuperaEvalConsumoReferencias(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaEvalConsumoReferencias " & pnCodEval)
        RecuperaEvalConsumoReferencias = D
        D = Nothing
    End Function

    Public Function RecuperaEvalCredDiario_CalcVentas(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaEvalCredDiario_CalcVentas " & pnCodEval)
        RecuperaEvalCredDiario_CalcVentas = D
        D = Nothing
    End Function

    Public Function RecuperaEvalCredDiario_Endeuda(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaEvalCredDiario_Endeuda " & pnCodEval)
        RecuperaEvalCredDiario_Endeuda = D
        D = Nothing
    End Function

    Public Sub RegistraEvalConsumo(ByVal D As ComTipos.clsDataset, ByVal dRefPers As ComTipos.clsDataset, ByVal dRefLab As ComTipos.clsDataset, ByVal dDeudas As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Eval_PersEvalConsumo_Grabar", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Eval_PersEvalConsumo_Grabar")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = D.ObtenerDato("nCodEval")
        oComm.Parameters.Add("@cCargo", SqlDbType.VarChar, 250).Value = D.ObtenerDato("cCargo")
        oComm.Parameters.Add("@cCentrotrabajo", SqlDbType.VarChar, 250).Value = D.ObtenerDato("cCentrotrabajo")
        oComm.Parameters.Add("@nCalificacion", SqlDbType.Int).Value = D.ObtenerDato("nCalificacion")
        oComm.Parameters.Add("@cAntiguedad", SqlDbType.VarChar, 50).Value = D.ObtenerDato("cAntiguedad")
        oComm.Parameters.Add("@cTiempolab", SqlDbType.VarChar, 60).Value = D.ObtenerDato("cTiempolab")
        oComm.Parameters.Add("@cComentarios", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComentarios")
        oComm.Parameters.Add("@nIngSol", SqlDbType.Money).Value = IIf(D.ObtenerDato("nIngSol") = "@#err", 0, D.ObtenerDato("nIngSol"))
        oComm.Parameters.Add("@nIngConyu", SqlDbType.Money).Value = IIf(D.ObtenerDato("nIngConyu") = "@#err", 0, D.ObtenerDato("nIngConyu"))
        oComm.Parameters.Add("@nRemExt", SqlDbType.Money).Value = IIf(D.ObtenerDato("nRemExt") = "@#err", 0, D.ObtenerDato("nRemExt"))
        oComm.Parameters.Add("@nOtrosIng", SqlDbType.Money).Value = IIf(D.ObtenerDato("nOtrosIng") = "@#err", 0, D.ObtenerDato("nOtrosIng"))
        oComm.Parameters.Add("@Ref_Pers", SqlDbType.Xml).Value = IIf(dRefPers.DevuelveXML = "@#err", "", dRefPers.DevuelveXML)
        oComm.Parameters.Add("@Ref_Lab", SqlDbType.Xml).Value = IIf(dRefLab.DevuelveXML = "@#err", "", dRefLab.DevuelveXML)
        oComm.Parameters.Add("@Deudas", SqlDbType.Xml).Value = IIf(dDeudas.DevuelveXML = "@#err", "", dDeudas.DevuelveXML)

        If CInt(D.ObtenerDato("nTipoGrati")) > 0 Then oComm.Parameters.Add("@nTipoGrati", SqlDbType.Int).Value = D.ObtenerDato("nTipoGrati") 'arcv 13-09-2010

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Sub RegistraEvalCredDiario(ByVal D As ComTipos.clsDataset, ByVal dCalcVentas As ComTipos.clsDataset, ByVal dEndeuda As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Eval_PersEvalCredDiario_Grabar", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Eval_PersEvalCredDiario_Grabar")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = D.ObtenerDato("nCodEval")
        oComm.Parameters.Add("@cComentarios", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComentarios")
        oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = IIf(D.ObtenerDato("nDisponible") = "@#err", 0, D.ObtenerDato("nDisponible"))
        oComm.Parameters.Add("@nCtasxCob", SqlDbType.Money).Value = IIf(D.ObtenerDato("nCtasxCob") = "@#err", 0, D.ObtenerDato("nCtasxCob"))
        oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = IIf(D.ObtenerDato("nInventario") = "@#err", 0, D.ObtenerDato("nInventario"))
        oComm.Parameters.Add("@nActFijo", SqlDbType.Money).Value = IIf(D.ObtenerDato("nActFijo") = "@#err", 0, D.ObtenerDato("nActFijo"))
        oComm.Parameters.Add("@Calc_Ventas", SqlDbType.Xml).Value = IIf(dCalcVentas.DevuelveXML = "@#err", "", dCalcVentas.DevuelveXML)
        oComm.Parameters.Add("@Nivel_Endeuda", SqlDbType.Xml).Value = IIf(dEndeuda.DevuelveXML = "@#err", "", dEndeuda.DevuelveXML)

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RegeneraEvaluacion(ByVal pnCodCred As Integer, _
                        ByVal pnCodAge As Integer, _
                        ByRef pnCodEvalNew As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("PersEval_RegeneraEvaluacion", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("PersEval_RegeneraEvaluacion")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nCodEvalNew", SqlDbType.Int).Direction = ParameterDirection.Output

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        pnCodEvalNew = oComm.Parameters.Item("@nCodEvalNew").Value

        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Sub PersEval_GrabaFotosDomi(ByVal nCodEval As Integer, ByVal oFoto() As Byte, ByVal nCodFoto As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("PersEval_RegeneraEvaluacion", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_GuardaFotosDomi ")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@oFoto", SqlDbType.Image).Value = oFoto
        oComm.Parameters.Add("@nCodFoto", SqlDbType.Int).Value = nCodFoto


        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Function RecuperaVerificacionDomiConsultaCreditoImagenes(ByVal nTipo As Integer, _
                        ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsEvalDomicFotos

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        Dim X As ComTipos.clsEvalDomicFotos
        X = New ComTipos.clsEvalDomicFotos
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_RecuperaImagenesDomicilio", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodTipo", SqlDbType.Int).Value = nTipo
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        R = oComm.ExecuteReader

        While R.Read

            X.oFoto = IIf(IsDBNull(R("oFoto")), Nothing, R("oFoto"))

        End While
        objCon = Nothing
        oComm = Nothing
        RecuperaVerificacionDomiConsultaCreditoImagenes = X

    End Function
    Public Sub RecuperaVerificacionDomiConsultaCredito(ByRef oDatos As ComTipos.clsPersEvalVerifDomic, _
                        ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer, _
                        Optional ByVal bAval As Boolean = False)
        'Eval_RecuperaVerificacionDomi

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        objCon = New ComConecta.clsConecta
        If bAval = False Then
            oComm = New SqlCommand("Eval_RecuperaVerificacionDomi_ConsultaCredito", objCon.DameConexion)
        Else
            oComm = New SqlCommand("Eval_RecuperaVerificacionAval_ConsultaCredito", objCon.DameConexion)
        End If
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        R = oComm.ExecuteReader
        'oComm.ExecuteNonQuery()

        While R.Read
            With oDatos
                .sReferenciaDomi = R("cReferenciaDomi")
                .sReferenciaNego = R("cReferenciaNego")
                .sReferenciaDomi2 = R("cReferenciaDomi2")
                .sReferenciaNego2 = R("cReferenciaNego2")

                If bAval Then
                    .sNomAval = R("cNomAval")
                    .sTelefAval = R("cTelefAval")
                    .sDirecAval = R("cDirecAval")
                End If

                .sReferenciaNego3 = R("cReferenciaNego3")
                .sReferenciaNego4 = R("cReferenciaNego4")
                .sReferenciaSumin = R("cReferenciaSumin")
                '-------------
            End With
        End While
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersEvalRecuperaReferenciasDomicCloud(ByRef oDatos As ComTipos.clsPersEvalVerifDomic, _
                      ByVal nCodEval As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        objCon = New ComConecta.clsConecta

        oComm = New SqlCommand("Eval_RecuperaReferenciasDomic ", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        R = oComm.ExecuteReader
        'oComm.ExecuteNonQuery()

        While R.Read
            With oDatos
                .dFecReg = R("dFecReg")
                .nCodEval = R("nCodEval")
                .sReferenciaDomi = R("cReferenciaDomi")
                .sReferenciaNego = R("cReferenciaNego")
                .sReferenciaDomi2 = R("cReferenciaDomi2")
                .sReferenciaNego2 = R("cReferenciaNego2")
                .sReferenciaNego3 = R("cReferenciaNego3")
                .sReferenciaNego4 = R("cReferenciaNego4")
                .sReferenciaSumin = R("cReferenciaSumin")
                '-------------
            End With
        End While
        objCon = Nothing
        oComm = Nothing
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''ivillafranca 14-05-2012'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function VerificaCondicionFlujo(ByVal nCodEval As Integer) As ComTipos.clsDataset
        Dim DatosCondicionFlujo As ComTipos.clsDataset
        DatosCondicionFlujo = New ComTipos.clsDataset
        DatosCondicionFlujo.CargaTabla(, , "VerificaCondicionFlujo " & nCodEval)
        VerificaCondicionFlujo = DatosCondicionFlujo
        DatosCondicionFlujo = Nothing
    End Function
    Public Function RecuperaDatosEntradaVentas(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsNuevoFlujoCaja
        Dim DataVentas As ComTipos.clsNuevoFlujoCaja
        DataVentas = New ComTipos.clsNuevoFlujoCaja
        'DataVentas.oCampañas.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 4044")
        DataVentas.oMesTipo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5812")
        DataVentas.oPeriodo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        DataVentas.oPeriodo1.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        DataVentas.oPeriodo2.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        DataVentas.oActivoFijo.CargaTabla(, , "FC_ActivoFijo")
        DataVentas.oVentasActuales.CargaTabla(, , "FC_RecuperaVentasActuales ")
        DataVentas.oVentasActualesDetalle.CargaTabla(, , "FC_RecuperaVentasActuales_Detalle")
        DataVentas.oCampañas.CargaTabla(, , "FC_RecuperaCampañas_Detalle")
        DataVentas.oSupuestos.CargaTabla(, , "FC_RecuperaSupuestos")
        DataVentas.oMesesAltos.CargaTabla(, , "FC_MesesAltos")
        DataVentas.oPorcentajeMaximo.CargaTabla(, , "FC_MontoMaximoProcentajedeVentas")
        DataVentas.DatosProceso.CargaTabla(, , "FC_RecuperaDatosProceso " & nCodEval & "," & nCodAge & "," & nCodCred)
        DataVentas.oDataCredito.CargaTabla(, , "DataCredito " & nCodEval & "," & nCodAge & "," & nCodCred)
        ' DataVentas.oCredAnt.CargaTabla(, , ""& nCodAge & "," & nCodCred)
        DataVentas.oCredAgricola.CargaTabla(, , "FC_RecuperaxCredAgricola ")
        '===========================
        'vvillafranca  19-07-2012 Nuevo Flujo de Caja
        DataVentas.oProyVentas.CargaTabla(, , "FC_ListaProyeccionVentasV " & nCodEval)
        DataVentas.oMesesSupuestos.CargaTabla(, , "FC_ListarMeses ")
        DataVentas.oIngresosExtra.CargaTabla(, , "FC_ListaIngresoExtra " & nCodEval)
        DataVentas.oListaOtrosCreditos.CargaTabla(, , "Pers_ListarOtrosCreditosFC " & nCodEval & "," & nCodAge & "," & nCodCred)
        '===========================

        RecuperaDatosEntradaVentas = DataVentas
        DataVentas = Nothing
    End Function
    Public Function RecuperaIngresosSalidasDEstinoModalidad(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer, Optional ByVal nCodPers As Integer = 0) As ComTipos.clsNuevoFlujoCaja
        Dim Datas As ComTipos.clsNuevoFlujoCaja
        Datas = New ComTipos.clsNuevoFlujoCaja
        Datas.oIngresoActivo.CargaTabla(, , "FC_RecuperaActivoFijo " & nCodEval & "," & nCodAge & "," & nCodCred)
        Datas.oIngresoCapital.CargaTabla(, , "FC_CT " & nCodEval & "," & nCodAge & "," & nCodCred)
        Datas.oIngresoMixto.CargaTabla(, , "FC_Mixto " & nCodEval & "," & nCodAge & "," & nCodCred)

        Datas.oEgresoActivo.CargaTabla(, , "FC_RecuperaActivoFijoEgreso " & nCodEval & "," & nCodAge & "," & nCodCred)
        Datas.oEgresoCapital.CargaTabla(, , "FC_CTEgreso " & nCodEval & "," & nCodAge & "," & nCodCred)
        Datas.oEgresoMixto.CargaTabla(, , "FC_MixtoEgreso " & nCodEval & "," & nCodAge & "," & nCodCred)
        ''''''' Datas.oCredAnt.CargaTabla(, , "RecuperaCreditosPendientes " & nCodPers)
        RecuperaIngresosSalidasDEstinoModalidad = Datas
        Datas = Nothing

    End Function
    Public Function RecuperaListaCreditosPendientes(ByVal nCodPers As Integer) As ComTipos.clsDataset
        Dim datos As ComTipos.clsDataset
        datos = New ComTipos.clsDataset
        datos.CargaTabla(, , "RecuperaCreditosPendientes " & nCodPers)
        RecuperaListaCreditosPendientes = datos
    End Function
    Public Function RecuperaDataAgricola(ByVal ncodAge As Integer, ByVal nCodCred As Integer, ByVal ncodEval As Integer) As ComTipos.clsDataset
        Dim dataagri As ComTipos.clsDataset
        dataagri = New ComTipos.clsDataset
        dataagri.CargaTabla(, , "FC_DevuelveAgriDatos " & ncodAge & "," & nCodCred & "," & ncodEval)
        RecuperaDataAgricola = dataagri
    End Function
    Public Sub GuardaNuevoFlujoCaja(ByVal oPV As ComTipos.clsDataset, _
                                    ByVal oRC As ComTipos.clsDataset, _
                                    ByVal oPN As ComTipos.clsDataset, _
                                    ByVal oIF As ComTipos.clsDataset, _
                                    ByVal oPF As ComTipos.clsDataset, _
                                    ByVal oUN As ComTipos.clsDataset, _
                                    ByVal oICT As ComTipos.clsDataset, _
                                    ByVal oIAF As ComTipos.clsDataset, _
                                    ByVal oIMX As ComTipos.clsDataset, _
                                    ByVal oECT As ComTipos.clsDataset, _
                                    ByVal oEAF As ComTipos.clsDataset, _
                                    ByVal oEMX As ComTipos.clsDataset, _
                                    ByVal oCamp As ComTipos.clsDataset, _
                                    ByVal oFinal As ComTipos.clsDataset, _
                                    ByVal nCodAge As Integer, _
                                    ByVal nCodCred As Integer, _
                                    ByVal nCodEval As Integer, _
                                    ByVal nCodPers As Integer, _
                                    ByRef ngraba As Integer)


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("GuardaNuevoFlujoCaja", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@oPV", SqlDbType.Xml)
        oComm.Parameters("@oPV").Value = oPV.DevuelveXML

        oComm.Parameters.Add("@oRC", SqlDbType.Xml)
        oComm.Parameters("@oRC").Value = oRC.DevuelveXML

        oComm.Parameters.Add("@oPN", SqlDbType.Xml)
        oComm.Parameters("@oPN").Value = oPN.DevuelveXML


        oComm.Parameters.Add("@oIF", SqlDbType.Xml)
        oComm.Parameters("@oIF").Value = oIF.DevuelveXML

        oComm.Parameters.Add("@oPF", SqlDbType.Xml)
        oComm.Parameters("@oPF").Value = oPF.DevuelveXML

        oComm.Parameters.Add("@oUN", SqlDbType.Xml)
        oComm.Parameters("@oUN").Value = oUN.DevuelveXML

        oComm.Parameters.Add("@oICT", SqlDbType.Xml)
        oComm.Parameters("@oICT").Value = oICT.DevuelveXML

        oComm.Parameters.Add("@oIAF", SqlDbType.Xml)
        oComm.Parameters("@oIAF").Value = oIAF.DevuelveXML

        oComm.Parameters.Add("@oIMX", SqlDbType.Xml)
        oComm.Parameters("@oIMX").Value = oIMX.DevuelveXML

        oComm.Parameters.Add("@oECT", SqlDbType.Xml)
        oComm.Parameters("@oECT").Value = oECT.DevuelveXML

        oComm.Parameters.Add("@oEAF", SqlDbType.Xml)
        oComm.Parameters("@oEAF").Value = oEAF.DevuelveXML

        oComm.Parameters.Add("@oEMX", SqlDbType.Xml)
        oComm.Parameters("@oEMX").Value = oEMX.DevuelveXML

        oComm.Parameters.Add("@oCamp", SqlDbType.Xml)
        oComm.Parameters("@oCamp").Value = oCamp.DevuelveXML

        oComm.Parameters.Add("@oFinal", SqlDbType.Xml)
        oComm.Parameters("@oFinal").Value = oFinal.DevuelveXML


        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers


        oComm.Parameters.Add("@ngraba", SqlDbType.Int).Value = ngraba
        oComm.Parameters.Item("@ngraba").Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        ngraba = oComm.Parameters.Item("@ngraba").Value
        oComm = Nothing
        objCon = Nothing

    End Sub
    Public Function RecuperaPeriodos() As ComTipos.ClsFlujoCajaSalida
        Dim Per As ComTipos.ClsFlujoCajaSalida
        Per = New ComTipos.ClsFlujoCajaSalida
        Per.oPeriodo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo1.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo2.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo3.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo4.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo6.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo7.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo8.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo9.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo10.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo11.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo12.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        Per.oPeriodo13.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        RecuperaPeriodos = Per
        Per = Nothing
    End Function
    Public Function RecuperaDataProcesoFlujo(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim Data As ComTipos.clsDataset
        Data = New ComTipos.clsDataset
        Data.CargaTabla(, , "FC_RecuperaDatosProceso " & nCodEval & "," & nCodAge & "," & nCodCred)
        RecuperaDataProcesoFlujo = Data
        Data = Nothing
    End Function
    Public Function RecuperaDataCredito(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim dat As ComTipos.clsDataset
        dat = New ComTipos.clsDataset
        dat.CargaTabla(, , "DataCredito " & nCodEval & "," & nCodAge & "," & nCodCred)
        RecuperaDataCredito = dat
    End Function
    Public Function RecuperaSalidaPeriodo() As ComTipos.clsDataset
        Dim Periodo As ComTipos.clsDataset
        Periodo = New ComTipos.clsDataset
        Periodo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 5814")
        RecuperaSalidaPeriodo = Periodo
        Periodo = Nothing
    End Function
    Public Function RecuperaFlujoDeBase(ByVal nCodAge As Integer, ByVal nCodCred As Integer, ByVal nCodEval As Integer) As ComTipos.clsDataset
        Dim data As ComTipos.clsDataset
        data = New ComTipos.clsDataset
        data.CargaTabla(, , "Rep_FC_Financiado " & nCodAge & "," & nCodCred & "," & nCodEval)
        RecuperaFlujoDeBase = data
        data = Nothing
    End Function

    Public Function VerificarGarantiaGrupo(ByVal nCodGrupo As Integer, ByVal nCiclo As Integer) As Boolean

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "sp_VerificarGarantiaPers " & nCodGrupo & "," & nCiclo)
        VerificarGarantiaGrupo = IIf(D.nNumFilas = 0, True, False)
        D = Nothing

    End Function

    Public Function RecuperaProyeccionVentas() As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevuelvePorcentajesMeses 74")
        RecuperaProyeccionVentas = D
        D = Nothing

    End Function

    Public Sub GuardaNuevoFlujoCajaEntradaDatos( _
                                    ByVal nCodEval As Integer, _
                                    ByVal nMesAlto As Double, _
                                    ByVal nMesNormal As Double, _
                                    ByVal nMesBajo As Double, _
                                    ByVal nSup1 As Double, _
                                    ByVal nMesIniSup1 As Integer, _
                                    ByVal nSup2 As Double, _
                                    ByVal nMesIniSup2 As Integer, _
                                    ByVal nSup3 As Double, _
                                    ByVal nMesIniSup3 As Integer, _
                                    ByVal nSup4 As Double, _
                                    ByVal nMesIniSup4 As Integer, _
                                    ByVal nSup5 As Double, _
                                    ByVal nMesIniSup5 As Integer, _
                                    ByVal nSup6 As Double, _
                                    ByVal nMesIniSup6 As Integer, _
                                    ByVal nSup7 As Double, _
                                    ByVal nMesIniSup7 As Integer, _
                                    ByVal nGastoAlto1 As Double, _
                                    ByVal nGastoAlto2 As Double, _
                                    ByVal nGastoAlto3 As Double, _
                                    ByVal nCantConsumo As Integer, _
                                    ByVal nCantNegocio As Integer, _
                                    ByVal nCantHipotec As Integer, _
                                    ByVal nMontoActFijo As Double, _
                                    ByVal oProyVentas As ComTipos.clsDataset, _
                                    ByVal oCampania As ComTipos.clsDataset, _
                                    ByVal oIngresoExtra As ComTipos.clsDataset, _
                                    ByVal oOtrasInstituciones As ComTipos.clsDataset)


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Pers_InsertaFlujoEntrada", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nMesAlto", SqlDbType.Money).Value = nMesAlto
        oComm.Parameters.Add("@nMesNormal", SqlDbType.Money).Value = nMesNormal
        oComm.Parameters.Add("@nMesBajo", SqlDbType.Money).Value = nMesBajo
        oComm.Parameters.Add("@nSup1", SqlDbType.Money).Value = nSup1
        oComm.Parameters.Add("@nMesIniSup1", SqlDbType.Int).Value = nMesIniSup1
        oComm.Parameters.Add("@nSup2", SqlDbType.Money).Value = nSup2
        oComm.Parameters.Add("@nMesIniSup2", SqlDbType.Int).Value = nMesIniSup2
        oComm.Parameters.Add("@nSup3", SqlDbType.Money).Value = nSup3
        oComm.Parameters.Add("@nMesIniSup3", SqlDbType.Int).Value = nMesIniSup3
        oComm.Parameters.Add("@nSup4", SqlDbType.Money).Value = nSup4
        oComm.Parameters.Add("@nMesIniSup4", SqlDbType.Int).Value = nMesIniSup4
        oComm.Parameters.Add("@nSup5", SqlDbType.Money).Value = nSup5
        oComm.Parameters.Add("@nMesIniSup5", SqlDbType.Int).Value = nMesIniSup5
        oComm.Parameters.Add("@nSup6", SqlDbType.Money).Value = nSup6
        oComm.Parameters.Add("@nMesIniSup6", SqlDbType.Int).Value = nMesIniSup6
        oComm.Parameters.Add("@nSup7", SqlDbType.Money).Value = nSup7
        oComm.Parameters.Add("@nMesIniSup7", SqlDbType.Int).Value = nMesIniSup7
        oComm.Parameters.Add("@nGastoAlto1", SqlDbType.Money).Value = nGastoAlto1
        oComm.Parameters.Add("@nGastoAlto2", SqlDbType.Money).Value = nGastoAlto2
        oComm.Parameters.Add("@nGastoAlto3", SqlDbType.Money).Value = nGastoAlto3
        oComm.Parameters.Add("@nCantConsumo", SqlDbType.Int).Value = nCantConsumo
        oComm.Parameters.Add("@nCantNegocio", SqlDbType.Int).Value = nCantNegocio
        oComm.Parameters.Add("@nCantHipotec", SqlDbType.Int).Value = nCantHipotec
        oComm.Parameters.Add("@nMontoActFijo", SqlDbType.Money).Value = nMontoActFijo

        oComm.Parameters.Add("@oProyVentas", SqlDbType.Xml)
        oComm.Parameters("@oProyVentas").Value = oProyVentas.DevuelveXML

        oComm.Parameters.Add("@oCampania", SqlDbType.Xml)
        oComm.Parameters("@oCampania").Value = oCampania.DevuelveXML

        oComm.Parameters.Add("@oIngresoExtra", SqlDbType.Xml)
        oComm.Parameters("@oIngresoExtra").Value = oIngresoExtra.DevuelveXML


        oComm.Parameters.Add("@oOtrasInstituciones", SqlDbType.Xml)
        oComm.Parameters("@oOtrasInstituciones").Value = oOtrasInstituciones.DevuelveXML


        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    'Public Function ListaDatosEntradaFC(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    D.CargaTabla(, , "Pers_ListaDatosEntradaFC " & nCodEval & "," & nCodAge & "," & nCodCred)
    '    ListaDatosEntradaFC = D
    '    D = Nothing

    'End Function

    'Public Function ListaProyEntradaFC(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    D.CargaTabla(, , "Pers_ListarProyVentasFC " & nCodEval & "," & nCodAge & "," & nCodCred)
    '    ListaProyEntradaFC = D
    '    D = Nothing

    'End Function

    'Public Function DevuelveFechaRegFC(ByVal nCodEval As Integer) As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    D.CargaTabla(, , "Pers_DevuelveFecRegFC " & nCodEval)
    '    DevuelveFechaRegFC = D
    '    D = Nothing

    'End Function

    'Public Function ListaCampaniaEntradaFC(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    D.CargaTabla(, , "Pers_ListarCampaniasFC " & nCodEval & "," & nCodAge & "," & nCodCred)
    '    ListaCampaniaEntradaFC = D
    '    D = Nothing

    'End Function

    'Public Function ListaIngreExtraFC(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    D.CargaTabla(, , "Pers_ListarIngreExtraFC " & nCodEval & "," & nCodAge & "," & nCodCred)
    '    ListaIngreExtraFC = D
    '    D = Nothing

    'End Function

    'Public Function ListaOtrosCreditos(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
    '    Dim D As ComTipos.clsDataset
    '    D = New ComTipos.clsDataset
    '    D.CargaTabla(, , "Pers_ListarOtrosCreditosFC " & nCodEval & "," & nCodAge & "," & nCodCred)
    '    ListaOtrosCreditos = D
    '    D = Nothing

    'End Function

    Public Function ListaDatosEntradaFC(ByVal nCodEval As Integer, ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsNuevoFlujoCaja
        Dim DataVentas As ComTipos.clsNuevoFlujoCaja
        DataVentas = New ComTipos.clsNuevoFlujoCaja

        DataVentas.oDatosEntradaFC.CargaTabla(, , "Pers_ListaDatosEntradaFC " & nCodEval & "," & nCodAge & "," & nCodCred)
        DataVentas.oListaProyEntradaFC.CargaTabla(, , "Pers_ListarProyVentasFC " & nCodEval & "," & nCodAge & "," & nCodCred)
        DataVentas.oFechaRegFC.CargaTabla(, , "Pers_DevuelveFecRegFC " & nCodEval)
        DataVentas.oListaCampaniaFC.CargaTabla(, , "Pers_ListarCampaniasFC " & nCodEval & "," & nCodAge & "," & nCodCred)
        DataVentas.oListaIngreExtraFC.CargaTabla(, , "Pers_ListarIngreExtraFC " & nCodEval & "," & nCodAge & "," & nCodCred)
        DataVentas.oListaOtrosCreditos.CargaTabla(, , "Pers_ListarOtrosCreditosFC " & nCodEval & "," & nCodAge & "," & nCodCred)

        ListaDatosEntradaFC = DataVentas
        DataVentas = Nothing
    End Function


    Public Function DevuelveFechaFlujo(ByVal nCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_FechaFlujoCaja " & nCodEval)
        DevuelveFechaFlujo = D
        D = Nothing

    End Function

    Public Function ValidaFlujoAntiguo(ByVal nCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "PersSelectorFlujoCaja " & nCodEval)
        ValidaFlujoAntiguo = D
        D = Nothing

    End Function

    Public Function RecuperaFotosEvaluacion(ByVal nCodEval As Long) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "PA_RecuperaFotosEvaluacion " & nCodEval)
        RecuperaFotosEvaluacion = D
        D = Nothing


    End Function
    '

    Public Function RecuperaFotoDeEvaluacion(ByVal nCodEval As Long, ByVal nIDFoto As Integer) As Byte()

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("RecuperaFotoEvaluacion", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nIDFoto", SqlDbType.Int).Value = nIDFoto

        R = oComm.ExecuteReader

        While R.Read

            RecuperaFotoDeEvaluacion = IIf(IsDBNull(R("iFoto")), Nothing, R("iFoto"))

        End While

        objCon = Nothing
        oComm = Nothing


    End Function
    Public Function ValidaFlujoAgencia(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_ValidaHabFlujoCaja " & nCodAge)
        ValidaFlujoAgencia = D
        D = Nothing

    End Function
    ''-------------------------------VVILLAFRANCA 21-08-2012  NUEVOS MANTENIMIENTOS EXPEDIENTE VIRTUAl -----------------------

    Public Sub Pers_EvalRegistraMargen(ByVal nCodEval As Integer, ByVal cProducto As String, ByVal nCostoUnit As Double, ByVal nVentasUnit As Double)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_EvalRegistraMargen", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("Cred_GuardaFotosDomi ")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@cProducto", SqlDbType.VarChar, 50).Value = cProducto
        oComm.Parameters.Add("@nCostoUnit", SqlDbType.Money).Value = nCostoUnit
        oComm.Parameters.Add("@nVentasUni", SqlDbType.Money).Value = nVentasUnit


        oComm.ExecuteNonQuery()
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Pers_EvalEliminaMargen(ByVal nCodEval As Integer, ByVal nItem As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_EvalEliminaMargen", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = nItem

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Pers_EvalRegistraMatPrima(ByVal nCodEval As Integer, ByVal nNumHoja As Integer, ByVal cMatPrim As String, ByVal cUniMed As String, ByVal nCostoUnit As Double, ByVal nCantidad As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_EvalRegistraMatPrima", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nNumHoja", SqlDbType.Int).Value = nNumHoja
        oComm.Parameters.Add("@cMatPrim", SqlDbType.VarChar, 500).Value = cMatPrim
        oComm.Parameters.Add("@cUniMed", SqlDbType.VarChar, 50).Value = cUniMed
        oComm.Parameters.Add("@nCostoUni", SqlDbType.Money).Value = nCostoUnit
        oComm.Parameters.Add("@nCantidad", SqlDbType.Money).Value = nCantidad

        oComm.ExecuteNonQuery()
   
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Pers_EvalRegistraInsumo(ByVal nCodEval As Integer, ByVal nNumHoja As Integer, ByVal cInsumo As String, ByVal nMonto As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_EvalRegistraInsumo", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nNumHoja", SqlDbType.Int).Value = nNumHoja
        oComm.Parameters.Add("@cInsumo", SqlDbType.VarChar, 500).Value = cInsumo
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Sub Pers_EvalEliminaMatPrima(ByVal nCodEval As Integer, ByVal nNumHoja As Integer, ByVal cMatPrim As String, ByVal nCostoUnit As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_EvalEliminaMatPrima", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nNumHoja", SqlDbType.Int).Value = nNumHoja
        oComm.Parameters.Add("@cMatPrim", SqlDbType.VarChar, 500).Value = cMatPrim
        oComm.Parameters.Add("@nCostoUni", SqlDbType.Money).Value = nCostoUnit

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub Pers_EvalEliminaInsumo(ByVal nCodEval As Integer, ByVal nNumHoja As Integer, ByVal cInsumo As String, ByVal nMonto As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_EvalEliminaInsumo", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nNumHoja", SqlDbType.Int).Value = nNumHoja
        oComm.Parameters.Add("@cInsumo", SqlDbType.VarChar, 500).Value = cInsumo
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    ''------------------------------------------------------------------------------------------------------------------------

    Public Function DevuelveTiempoCargaFotos(ByVal nTamanio As Double) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Eval_CredDevuelvetiempo " & nTamanio)
        DevuelveTiempoCargaFotos = D
        D = Nothing
    End Function

    Public Function RecuperaNivelVentas(ByVal nCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_RecuperaNivelVentas " & nCodEval)
        RecuperaNivelVentas = D
        D = Nothing
    End Function

    Public Sub GuardaNiveldeVentas(ByVal nCodEval As Integer, ByVal oMeses As ComTipos.clsDataset, _
                                    ByVal oDias As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Eval_RegistraNivelVentas", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval

        oComm.Parameters.Add("@oNivelVentas", SqlDbType.Xml)
        oComm.Parameters("@oNivelVentas").Value = oMeses.DevuelveXML

        oComm.Parameters.Add("@oNivelVentas2", SqlDbType.Xml)
        oComm.Parameters("@oNivelVentas2").Value = oDias.DevuelveXML

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function CargarDatosExpedienteVirtual(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_DatosExpVirtual " & nCodAge & "," & nCodCred)
        CargarDatosExpedienteVirtual = D
        D = Nothing
    End Function

    Public Function ValidaNuevaEvaluacion(ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_ValidaHabEvaluacion " & nCodAge)
        ValidaNuevaEvaluacion = D
        D = Nothing

    End Function

    Public Sub PersEvalRegistraExpedienteVirtual(ByVal nCodAge As Integer, _
                          ByVal nCodCred As Integer, ByVal cCodUsu As String, ByVal oExpediente() As Byte, ByVal Anexo1() As Byte, ByVal iFotos() As Byte)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraExpedienteVirtual")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
        oComm.Parameters.Add("@iExpediente", SqlDbType.Image).Value = IIf(oExpediente Is Nothing, DBNull.Value, oExpediente)
        oComm.Parameters.Add("@iAnexoI", SqlDbType.Image).Value = IIf(Anexo1 Is Nothing, DBNull.Value, Anexo1)
        oComm.Parameters.Add("@iFotos", SqlDbType.Image).Value = IIf(iFotos Is Nothing, DBNull.Value, iFotos)

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Function RecuperaExpedienteVirtual(ByVal nCodAge As Integer, ByVal nCodCred As Integer) As ComTipos.clsEvalDomicFotos

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader
        Dim X As ComTipos.clsEvalDomicFotos
        X = New ComTipos.clsEvalDomicFotos

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_DevuelveExpedienteVirtual", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        R = oComm.ExecuteReader
        While R.Read
            X.oExpediente = IIf(IsDBNull(R("iExpediente")), Nothing, R("iExpediente"))
            X.oAnexoI = IIf(IsDBNull(R("iAnexo1")), Nothing, R("iAnexo1"))
            X.oAnexoFotos = IIf(IsDBNull(R("iFotos")), Nothing, R("iFotos"))
        End While
        objCon = Nothing
        oComm = Nothing
        RecuperaExpedienteVirtual = X

    End Function

    Public Function RecuperaNuevaVerificacionDomi(ByVal nCodEval As Integer, _
                        Optional ByVal bAval As Boolean = False) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset

        If bAval = False Then
            D.CargaTabla(, , "Eval_RecuperaNuevaVerificacionDomi " & nCodEval)
        Else
            D.CargaTabla(, , "Eval_RecuperaNuevaVerificacionDomi " & nCodEval)
        End If
        RecuperaNuevaVerificacionDomi = D
        D = Nothing
    End Function

    Public Function RecuperaListadoDocumentos_Eval(ByVal nCodAge As Integer, _
                       ByVal nCodCred As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Pers_ListaDocumentosxCredito " & nCodAge & " , " & nCodCred)
        RecuperaListadoDocumentos_Eval = D
        D = Nothing
    End Function

    Public Sub PersRegistraDocumento(ByVal nCodAge As Integer, ByVal nCodCred As Integer, _
                           ByVal oArchivo() As Byte, ByVal cNomArchi As String, ByVal cExtension As String, ByVal cCodUsu As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_Inserta_Documento", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("Pers_Inserta_Documento")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@oArchivo", SqlDbType.Image).Value = IIf(oArchivo Is Nothing, DBNull.Value, oArchivo)
        oComm.Parameters.Add("@cNomArchivo", SqlDbType.VarChar, 200).Value = cNomArchi
        oComm.Parameters.Add("@cExtension", SqlDbType.VarChar, 10).Value = cExtension
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 200).Value = cCodUsu

        '---------------

        'ARCV 11-05-2011
        oComm.ExecuteNonQuery()
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function RecuperaDocumento_Sustento(ByVal nCodAge As Integer, ByVal nCodCred As Integer, ByVal nItem As Integer) As ComTipos.clsEvalDomicFotos

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim R As SqlDataReader

        Dim X As ComTipos.clsEvalDomicFotos
        X = New ComTipos.clsEvalDomicFotos
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_DevuelveDocumento", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = nItem
        R = oComm.ExecuteReader

        While R.Read

            X.oFoto = IIf(IsDBNull(R("oArchivo")), Nothing, R("oArchivo"))

        End While
        objCon = Nothing
        oComm = Nothing
        RecuperaDocumento_Sustento = X

    End Function

    Public Sub PersEliminaDocumento(ByVal nCodAge As Integer, ByVal nCodCred As Integer, ByVal nItem As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_Elimina_Documento", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("Pers_Inserta_Documento")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = nItem
        '---------------

        'ARCV 11-05-2011
        oComm.ExecuteNonQuery()
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function Pers_ValidaDocumentos(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim J As ComTipos.clsDataset
        J = New ComTipos.clsDataset
        J.CargaTabla(, , "Pers_ValidaCredito " & pnCodAge & "," & pnCodCred)
        Pers_ValidaDocumentos = J
        J = Nothing
    End Function

    Public Sub PersEvalBalGen_Grabar_New(ByVal nCodEval As Integer, _
                       ByVal nDisponible As Double, _
                       ByVal nAdelantoProv As Double, _
                       ByVal nCtasCobrar As Double, _
                       ByVal nInventario As Double, _
                       ByVal nMercaderiaTran As Double, _
                       ByVal nBienesMueb As Double, _
                       ByVal nBienesInmue As Double, _
                       ByVal nCtasPagar As Double, _
                       ByVal nAdelantoCli As Double, _
                       ByVal nOtrasCtasPag As Double, _
                       ByVal nPasivoNC As Double, _
                       ByVal nCtasPagarLP As Double, _
                       ByVal nOtrasCPagarLP As Double, _
                       ByVal nCapital As Double, _
                       ByVal nReservas As Double, _
                       ByVal nDistribUtilidad As Double, _
                       ByVal nMateriales As Double, _
                       ByVal nMatPrima As Double, _
                       ByVal nMatProceso As Double, _
                       ByVal nProdTerminado As Double, _
                       Optional ByVal nInversiones As Double = 0, _
                       Optional ByVal nDeudaActivoF As Double = 0, _
                       Optional ByVal bInversiones As Boolean = False, _
                       Optional ByVal bMuebles As Boolean = False, _
                       Optional ByVal bInMuebles As Boolean = False)


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Eval_PersEvalBalGen_Grabar_NEW", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = nDisponible
        oComm.Parameters.Add("@nAdelantoProv", SqlDbType.Money).Value = nAdelantoProv
        oComm.Parameters.Add("@nCtasCobrar", SqlDbType.Money).Value = nCtasCobrar
        oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = nInventario
        oComm.Parameters.Add("@nMercaderiaTran", SqlDbType.Money).Value = nMercaderiaTran
        oComm.Parameters.Add("@nBienesMueb", SqlDbType.Money).Value = nBienesMueb
        oComm.Parameters.Add("@nBienesInmue", SqlDbType.Money).Value = nBienesInmue
        oComm.Parameters.Add("@nCtasPagar", SqlDbType.Money).Value = nCtasPagar
        oComm.Parameters.Add("@nAdelantoCli", SqlDbType.Money).Value = nAdelantoCli
        oComm.Parameters.Add("@nOtrasCtasPag", SqlDbType.Money).Value = nOtrasCtasPag
        oComm.Parameters.Add("@nPasivoNoC", SqlDbType.Money).Value = nPasivoNC

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

        oComm.Parameters.Add("@bInversiones", SqlDbType.Bit).Value = bInversiones
        oComm.Parameters.Add("@bMuebles", SqlDbType.Bit).Value = bMuebles
        oComm.Parameters.Add("@bInmuebles", SqlDbType.Bit).Value = bInMuebles

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function Pers_ValidaFormatosEvaluacion(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As ComTipos.clsDataset
        Dim J As ComTipos.clsDataset
        J = New ComTipos.clsDataset
        J.CargaTabla(, , "Pers_ValidaDatosAntiguosEval " & pnCodAge & "," & pnCodCred)
        Pers_ValidaFormatosEvaluacion = J
        J = Nothing
    End Function

    Public Sub InsertaDatos_EvalAgricola(ByVal nCodEval As Integer, ByVal cCultivo As String, ByVal nVariedad As Integer, ByVal dFecSiembraI As Date _
    , ByVal dFecSiembraT As Date, ByVal dFecCosechaI As Date, ByVal dFecCosechaT As Date, ByVal nPeriodoVeg As Integer, ByVal nNumHasAFinan As Double _
    , ByVal nRendProm As Double, ByVal nRendSen As Double, ByVal nPrecioProm As Double, ByVal nPrecioSen As Double, ByVal nTipoTecnologia As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_InsertaDatosEvalAgricola", objCon.DameConexion)

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@cCultivo", SqlDbType.VarChar, 300).Value = cCultivo
        oComm.Parameters.Add("@nVariedad", SqlDbType.Int).Value = nVariedad
        oComm.Parameters.Add("@dFecSiembraI", SqlDbType.DateTime).Value = dFecSiembraI
        oComm.Parameters.Add("@dFecSiembraT", SqlDbType.DateTime).Value = dFecSiembraT
        oComm.Parameters.Add("@dFecCosechaI", SqlDbType.DateTime).Value = dFecCosechaI
        oComm.Parameters.Add("@dFecCosechaT", SqlDbType.DateTime).Value = dFecCosechaT
        oComm.Parameters.Add("@nPeriodoVeg", SqlDbType.Int).Value = nPeriodoVeg
        oComm.Parameters.Add("@nNumHasAFinan", SqlDbType.Money).Value = nNumHasAFinan
        oComm.Parameters.Add("@nRendProm", SqlDbType.Money).Value = nRendProm
        oComm.Parameters.Add("@nRendSen", SqlDbType.Money).Value = nRendSen
        oComm.Parameters.Add("@nPrecioProm", SqlDbType.Money).Value = nPrecioProm
        oComm.Parameters.Add("@nPrecioSen", SqlDbType.Money).Value = nPrecioSen
        oComm.Parameters.Add("@nTipoTecnologia", SqlDbType.Int).Value = nTipoTecnologia

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub
    Public Function CargaDatosEvalAgricola(ByVal nCodEval As Integer) As ComTipos.clsNewEvalAgricola
        Dim DatosAgri As ComTipos.clsNewEvalAgricola
        DatosAgri = New ComTipos.clsNewEvalAgricola

        DatosAgri.oTecnologia.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 67")
        DatosAgri.oVariedad.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 68")
        DatosAgri.oUniMedida.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 66")
        DatosAgri.oDatosCredAgri.CargaTabla(, , "PersAgri_Datos " & nCodEval)
        DatosAgri.oLaboresCultivo.CargaTabla(, , "PersAgri_ListaActividad " & nCodEval & ",1")
        DatosAgri.oMaquinaria.CargaTabla(, , "PersAgri_ListaActividad " & nCodEval & ",2")
        DatosAgri.oSemilla.CargaTabla(, , "PersAgri_ListaActividad " & nCodEval & ",3")
        DatosAgri.oFertilizantes.CargaTabla(, , "PersAgri_ListaActividad " & nCodEval & ",4")
        DatosAgri.oHerbicidas.CargaTabla(, , "PersAgri_ListaActividad " & nCodEval & ",5")
        DatosAgri.oAgua.CargaTabla(, , "PersAgri_ListaActividad " & nCodEval & ",6")
        DatosAgri.oCostosIndirectos.CargaTabla(, , "PersAgri_ListaActividad " & nCodEval & ",7")

        CargaDatosEvalAgricola = DatosAgri
        DatosAgri = Nothing
    End Function

    Public Sub GuardaNewEvalAgricola(ByVal oPV As ComTipos.clsDataset, _
                                    ByVal oRC As ComTipos.clsDataset, _
                                    ByVal oPN As ComTipos.clsDataset, _
                                    ByVal oIF As ComTipos.clsDataset, _
                                    ByVal oPF As ComTipos.clsDataset, _
                                    ByVal oUN As ComTipos.clsDataset, _
                                    ByVal oICT As ComTipos.clsDataset, _
                                     ByVal nCodEval As Integer)


        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Pers_GuardarCostosProduccionAgri", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@oLaboresCultivo", SqlDbType.Xml)
        oComm.Parameters("@oLaboresCultivo").Value = oPV.DevuelveXML

        oComm.Parameters.Add("@oMaquinaria", SqlDbType.Xml)
        oComm.Parameters("@oMaquinaria").Value = oRC.DevuelveXML

        oComm.Parameters.Add("@oSemilla", SqlDbType.Xml)
        oComm.Parameters("@oSemilla").Value = oPN.DevuelveXML


        oComm.Parameters.Add("@oFertilizantes", SqlDbType.Xml)
        oComm.Parameters("@oFertilizantes").Value = oIF.DevuelveXML

        oComm.Parameters.Add("@oHerbicidas", SqlDbType.Xml)
        oComm.Parameters("@oHerbicidas").Value = oPF.DevuelveXML

        oComm.Parameters.Add("@oAgua", SqlDbType.Xml)
        oComm.Parameters("@oAgua").Value = oUN.DevuelveXML

        oComm.Parameters.Add("@oCostosIndirectos", SqlDbType.Xml)
        oComm.Parameters("@oCostosIndirectos").Value = oICT.DevuelveXML

        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    ' vvillafranca 10-10-2012 ventas no dependan de la UF
    Public Sub RecuperaEstadGananPerd_Ventas(ByRef oDatos As ComTipos.clsDataset, _
     ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_RecuperaEstadGananPerd_ventas " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub

    Public Sub CalculaEstadGananPerd_Ventas(ByRef oDatos As ComTipos.clsDataset, _
ByVal nCodEval As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Eval_CalculaEstadGananPerd_Ventas " & nCodEval)
        oDatos = oData
        oData = Nothing
    End Sub
    ''''''


    Public Function CargaDatosEvalAgricola_Cronograma(ByVal nCodEval As Integer) As ComTipos.clsNewEvalAgricola
        Dim DatosAgri As ComTipos.clsNewEvalAgricola
        DatosAgri = New ComTipos.clsNewEvalAgricola

        DatosAgri.oDatosCredAgri.CargaTabla(, , "PersAgri_Datos " & nCodEval)
        DatosAgri.oLaboresCultivo.CargaTabla(, , "PersAgri_ListaActividad_Cronograma " & nCodEval & ",1")
        DatosAgri.oMaquinaria.CargaTabla(, , "PersAgri_ListaActividad_Cronograma " & nCodEval & ",2")
        DatosAgri.oSemilla.CargaTabla(, , "PersAgri_ListaActividad_Cronograma " & nCodEval & ",3")
        DatosAgri.oFertilizantes.CargaTabla(, , "PersAgri_ListaActividad_Cronograma " & nCodEval & ",4")
        DatosAgri.oHerbicidas.CargaTabla(, , "PersAgri_ListaActividad_Cronograma " & nCodEval & ",5")
        DatosAgri.oAgua.CargaTabla(, , "PersAgri_ListaActividad_Cronograma " & nCodEval & ",6")
        DatosAgri.oCostosIndirectos.CargaTabla(, , "PersAgri_ListaActividad_Cronograma " & nCodEval & ",7")

        CargaDatosEvalAgricola_Cronograma = DatosAgri
        DatosAgri = Nothing
    End Function

    Public Function Cred_DevuelveListaExpedientesxGenerar() As ComTipos.clsDataset
        Dim J As ComTipos.clsDataset
        J = New ComTipos.clsDataset
        J.CargaTabla(, , "Cred_DevuelveListaExpedientesxGenerar ")
        Cred_DevuelveListaExpedientesxGenerar = J
        J = Nothing
    End Function

    Public Sub PersActualizaGeneracionExpediente(ByVal nCodAge As Integer, ByVal nCodCred As Integer, ByVal nResult As Integer, ByVal nError As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_GrabaResultadoExpediente", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nResult", SqlDbType.Int).Value = nResult
        oComm.Parameters.Add("@nError", SqlDbType.Int).Value = nError

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub
End Class

