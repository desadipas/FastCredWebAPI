Imports System.Data
Imports System.Data.SqlClient

Public Class ClsPersona

    Dim oCons As ComConstante.clsConstante
    Dim fPersona As ComTipos.clsCargaPersona

    Public Function RecuperaCuotaFam(ByVal pnCodEval As Integer, ByVal pnTipoPers As Integer) As Double
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RecuperaCuotaFam ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@nCuotaFam", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nTipoPers", SqlDbType.Int).Value = pnTipoPers


        oComm.ExecuteNonQuery()
        RecuperaCuotaFam = oComm.Parameters(1).Value

        oComm = Nothing
        objCon = Nothing

    End Function

    Public Sub RegistroHojaCostos(ByVal pnCodEval As Integer, ByVal psProd1 As String, _
        ByVal psProd2 As String, ByVal psProd3 As String, ByVal psProd4 As String, ByVal psProd5 As String, ByVal psProd6 As String, _
        ByVal oDataMatPri As ComTipos.clsDataset, ByVal oDataMatPri2 As ComTipos.clsDataset, _
        ByVal oDataMatPri3 As ComTipos.clsDataset, ByVal oDataMatPri4 As ComTipos.clsDataset, ByVal oDataMatPri5 As ComTipos.clsDataset, ByVal oDataMatPri6 As ComTipos.clsDataset, _
        ByVal oDataInsumo As ComTipos.clsDataset, ByVal oDataInsumo2 As ComTipos.clsDataset, _
        ByVal oDataInsumo3 As ComTipos.clsDataset, ByVal oDataInsumo4 As ComTipos.clsDataset, ByVal oDataInsumo5 As ComTipos.clsDataset, ByVal oDataInsumo6 As ComTipos.clsDataset, ByVal pnCant1 As Integer, ByVal pnCant2 As Integer, ByVal pnCant3 As Integer, ByVal pnCant4 As Integer, ByVal pnCant5 As Integer, ByVal pnCant6 As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraHojaCosto ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@pxDatos", SqlDbType.Xml).Value = oDataMatPri.DevuelveXML
        oComm.Parameters.Add("@pxDatos2", SqlDbType.Xml).Value = oDataMatPri2.DevuelveXML
        oComm.Parameters.Add("@pxDatos3", SqlDbType.Xml).Value = oDataMatPri3.DevuelveXML
        oComm.Parameters.Add("@pxDatos4", SqlDbType.Xml).Value = oDataMatPri4.DevuelveXML
        oComm.Parameters.Add("@pxDatos5", SqlDbType.Xml).Value = oDataMatPri5.DevuelveXML
        oComm.Parameters.Add("@pxDatos6", SqlDbType.Xml).Value = oDataMatPri6.DevuelveXML
        oComm.Parameters.Add("@pxDatosInsumo", SqlDbType.Xml).Value = oDataInsumo.DevuelveXML
        oComm.Parameters.Add("@pxDatosInsumo2", SqlDbType.Xml).Value = oDataInsumo2.DevuelveXML
        oComm.Parameters.Add("@pxDatosInsumo3", SqlDbType.Xml).Value = oDataInsumo3.DevuelveXML
        oComm.Parameters.Add("@pxDatosInsumo4", SqlDbType.Xml).Value = oDataInsumo4.DevuelveXML
        oComm.Parameters.Add("@pxDatosInsumo5", SqlDbType.Xml).Value = oDataInsumo5.DevuelveXML
        oComm.Parameters.Add("@pxDatosInsumo6", SqlDbType.Xml).Value = oDataInsumo6.DevuelveXML
        oComm.Parameters.Add("@psProd1", SqlDbType.VarChar, 250).Value = psProd1
        oComm.Parameters.Add("@psProd2", SqlDbType.VarChar, 250).Value = psProd2
        oComm.Parameters.Add("@psProd3", SqlDbType.VarChar, 250).Value = psProd3
        oComm.Parameters.Add("@psProd4", SqlDbType.VarChar, 250).Value = psProd4
        oComm.Parameters.Add("@psProd5", SqlDbType.VarChar, 250).Value = psProd5
        oComm.Parameters.Add("@psProd6", SqlDbType.VarChar, 250).Value = psProd6
        oComm.Parameters.Add("@pnNumDatos1", SqlDbType.Int).Value = oDataMatPri.nNumFilas
        oComm.Parameters.Add("@pnNumDatos2", SqlDbType.Int).Value = oDataMatPri2.nNumFilas
        oComm.Parameters.Add("@pnNumDatos3", SqlDbType.Int).Value = oDataMatPri3.nNumFilas
        oComm.Parameters.Add("@pnNumDatos4", SqlDbType.Int).Value = oDataMatPri4.nNumFilas
        oComm.Parameters.Add("@pnNumDatos5", SqlDbType.Int).Value = oDataMatPri5.nNumFilas
        oComm.Parameters.Add("@pnNumDatos6", SqlDbType.Int).Value = oDataMatPri6.nNumFilas
        oComm.Parameters.Add("@pnNumInsu1", SqlDbType.Int).Value = oDataInsumo.nNumFilas
        oComm.Parameters.Add("@pnNumInsu2", SqlDbType.Int).Value = oDataInsumo2.nNumFilas
        oComm.Parameters.Add("@pnNumInsu3", SqlDbType.Int).Value = oDataInsumo3.nNumFilas
        oComm.Parameters.Add("@pnNumInsu4", SqlDbType.Int).Value = oDataInsumo4.nNumFilas
        oComm.Parameters.Add("@pnNumInsu5", SqlDbType.Int).Value = oDataInsumo5.nNumFilas
        oComm.Parameters.Add("@pnNumInsu6", SqlDbType.Int).Value = oDataInsumo6.nNumFilas
        oComm.Parameters.Add("@pnCant1", SqlDbType.Int).Value = pnCant1
        oComm.Parameters.Add("@pnCant2", SqlDbType.Int).Value = pnCant2
        oComm.Parameters.Add("@pnCant3", SqlDbType.Int).Value = pnCant3
        oComm.Parameters.Add("@pnCant4", SqlDbType.Int).Value = pnCant4
        oComm.Parameters.Add("@pnCant5", SqlDbType.Int).Value = pnCant5
        oComm.Parameters.Add("@pnCant6", SqlDbType.Int).Value = pnCant6


        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function BuscaHomonimos(ByVal psApePat As String, ByVal psApeMat As String, ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_BuscaHomonimos '" & psApePat & "','" & psApeMat & "'," & pnCodPers)
        BuscaHomonimos = D
        D = Nothing
    End Function

    Public Function BuscaHomonimosBaseNegativa(ByVal pnCodPers As Integer, ByVal pnCodAge As Integer, ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_BuscaHomonimosBaseNegativa " & pnCodPers & ", " & pnCodAge & ", '" & pcCodUsu & "'")
        BuscaHomonimosBaseNegativa = D
        D = Nothing
    End Function

    Public Function RUCExsiste(ByVal psRUC As String, ByVal pnCodPers As Integer) As Boolean
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_BuscaRUCDuplicado '" & psRUC & "'," & pnCodPers)
        If D.nNumFilas > 0 Then
            RUCExsiste = True
        Else
            RUCExsiste = False
        End If

        D = Nothing
    End Function

    Public Function DNIExsiste(ByVal psDNI As String, ByVal pnCodPers As Integer) As Boolean
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_BuscaDNIDuplicado '" & psDNI & "'," & pnCodPers)
        If D.nNumFilas > 0 Then
            DNIExsiste = True
        Else
            DNIExsiste = False
        End If

        D = Nothing
    End Function
    Public Sub RegistraComentario(ByVal pnCodPers As Integer, ByVal psComen As String)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_RegistraComentario", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@psComen", SqlDbType.Text).Value = psComen
        
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing
    End Sub
    Public Function RecuperaComentarios(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_RecuperaComentarios " & pnCodPers)
        RecuperaComentarios = D
        D = Nothing
    End Function

    Public Function RecuperaPosicionCliente(ByVal psDoc As String) As ComTipos.ClsPosicionCli
        Dim D As ComTipos.ClsPosicionCli
        Dim sCodEdu As String

        D = New ComTipos.ClsPosicionCli

        D.RCCTemp.CargaTabla(, , "CredW_RecuperaRCC '" & psDoc & "'")
        sCodEdu = D.RCCTemp.ObtenerDato("Cod_Edu")
        D.RCC.CargaTabla(, , "CredW_RecuperaRCCTotal '" & psDoc & "'")
        D.RCCDet.CargaTabla(, , "CredW_RecuperaRCCTotalDet '" + sCodEdu + "'")
        D.Datos.CargaTabla(, , " Cred_RecuperaPosicionCLi '" & psDoc & "'")
        D.Comen.CargaTabla(, , " Pers_RecuperaComentarios '" & D.Datos.ObtenerDato("nCodPers") & "'")
        D.BaseNegativa.CargaTabla(, , " Cred_DevuelveDatosPersonasBaseNegativaCentralRiesgo '" & psDoc & "'")
        D.TamizGenesys.CargaTabla(, , " Cred_DevuelveFiltroTAMIZ '" & psDoc & "'")
        D.SCORD.CargaTabla(, , " Cred_BuscarFiltroDATACREDITO_Historico '" & psDoc & "'")  'ARCV 28-02-2012
        RecuperaPosicionCliente = D
        D = Nothing
    End Function
    Public Function ActualizaPersonasDuplicados(ByVal pnCodPersBusqueda As Integer, ByVal pnCodPersReemplazo As Integer, ByVal cUser As String) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("ConsolidacionClientes", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodPersBus", SqlDbType.Int).Value = pnCodPersBusqueda
        oComm.Parameters.Add("@nCodPersReemplazo", SqlDbType.Int).Value = pnCodPersReemplazo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
        oComm.Parameters.Add("@pnResul", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        ActualizaPersonasDuplicados = oComm.Parameters("@pnResul").Value

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
    End Function
    Public Function RegistraPersona(ByVal pnCodPers As Integer, ByVal pnCodAge As Integer, ByVal poData As ComTipos.clsDataset, ByVal pDocs As ComTipos.clsDataset, ByVal pParen As ComTipos.clsDataset, ByVal pDepen As ComTipos.clsDataset, ByVal pSucursal As ComTipos.clsDataset, ByVal pRepresentantes As ComTipos.clsDataset, ByVal pAccionistas As ComTipos.clsDataset) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_RegistraPersona", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pData", SqlDbType.Xml).Value = poData.DevuelveXML
        oComm.Parameters.Add("@pDocs", SqlDbType.Xml).Value = pDocs.DevuelveXML
        oComm.Parameters.Add("@pParen", SqlDbType.Xml).Value = pParen.DevuelveXML
        oComm.Parameters.Add("@pDepen", SqlDbType.Xml).Value = pDepen.DevuelveXML
        oComm.Parameters.Add("@pSucursal", SqlDbType.Xml).Value = pSucursal.DevuelveXML
        oComm.Parameters.Add("@pRepresentantes", SqlDbType.Xml).Value = pRepresentantes.DevuelveXML
        oComm.Parameters.Add("@pAccionistas", SqlDbType.Xml).Value = pAccionistas.DevuelveXML
        oComm.Parameters.Add("@pnCodPersNew", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        RegistraPersona = oComm.Parameters("@pnCodPersNew").Value

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing

    End Function

    Public Function RecuperaDatosDirecc(ByVal pnCodPers As Integer) As ComTipos.ClsPersDireccion
        Dim D As ComTipos.ClsPersDireccion
        D = New ComTipos.ClsPersDireccion
        D.oDatosDir.CargaTabla(, , "Pers_RecuperaDireccion " & pnCodPers)
        RecuperaDatosDirecc = D

        D = Nothing
    End Function

    Public Function RecuperaDatosDireccTrab(ByVal pnCodPers As Integer) As ComTipos.ClsPersDireccion
        Dim D As ComTipos.ClsPersDireccion
        D = New ComTipos.ClsPersDireccion
        D.oDatosDir.CargaTabla(, , "Pers_RecuperaDireccionTrab " & pnCodPers)
        RecuperaDatosDireccTrab = D

        D = Nothing
    End Function

    Public Function RecuperaDatosNegocio(ByVal pnCodPers As Integer) As ComTipos.ClsPersDireccion
        Dim D As ComTipos.ClsPersDireccion
        D = New ComTipos.ClsPersDireccion
        D.oDatosDir.CargaTabla(, , "Pers_RecuperaInformacionNegocio " & pnCodPers)
        RecuperaDatosNegocio = D

        D = Nothing
    End Function

    Public Function RecuperaCIIU(ByVal psBusq As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_RecuperaCIIU '" & psBusq & "'")
        RecuperaCIIU = D
        D = Nothing
    End Function
    Public Function RecuperaTiposParentesco() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "CredW_RecuperaParentezcoMicroSeguro ")
        RecuperaTiposParentesco = D
        D = Nothing
    End Function

    Public Function RecuperaFiltros(ByVal psBusq As String, ByVal pnTipo As Integer, Optional ByVal pnCodAge As Integer = 0, Optional ByVal pnCodSubProd As Integer = 0) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_Filtros '" & psBusq & "', " & pnTipo & ", " & pnCodAge & ", " & pnCodSubProd)
        RecuperaFiltros = D
        D = Nothing
    End Function

    Public Function RecuperaNacionalidad(ByVal psBusq As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_RecuperaPaises '" & psBusq & "'")
        RecuperaNacionalidad = D
        D = Nothing
    End Function

    Public Function RecuperaDatosPersonaMicroSeguro(ByVal pnCodPers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_RecuperaDatosPersona " & pnCodPers)
        RecuperaDatosPersonaMicroSeguro = D
        D = Nothing
    End Function

    Public Function ActualizarDatosDEPersona(ByVal pnCodPers As Integer) As Boolean

        Dim oComm As SqlCommand
        Dim oCon As ComConecta.clsConecta

        oCon = New ComConecta.clsConecta

        oComm = New SqlCommand("Pers_SeDebeActualizarDatosDEPersona", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pbResul", SqlDbType.Bit).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        ActualizarDatosDEPersona = oComm.Parameters(1).Value

        oCon.CierraConexion()
        oComm = Nothing

    End Function

    Public Function RecuperaDatosPersona(ByVal pnCodPers As Integer) As ComTipos.ClsPersona

        Dim C As ComTipos.ClsPersona
        C = New ComTipos.ClsPersona

        C.Data.CargaTabla(, , "Pers_RecuperaDatosPersona " & pnCodPers)
        C.Documentos.CargaTabla(, , "Pers_RecuperaDocumentos " & pnCodPers)
        C.Parentesco.CargaTabla(, , "Pers_RecuperaParentescos " & pnCodPers)
        C.Dependientes.CargaTabla(, , "Pers_RecuperaDependientes " & pnCodPers)
        C.oSucursales.CargaTabla(, , "Pers_RecuperaSucursales " & pnCodPers)
        C.oRepresentantes.CargaTabla(, , "Pers_RecuperaRepresentantes " & pnCodPers)
        C.oAccionistas.CargaTabla(, , "Pers_RecuperaAccionistas " & pnCodPers)
        C.oDirec = RecuperaDatosDirecc(pnCodPers)
        C.oDirecTrab = RecuperaDatosDireccTrab(pnCodPers)
        C.oNegocio = RecuperaDatosNegocio(pnCodPers)


        RecuperaDatosPersona = C
        C = Nothing

    End Function
    Public Function RecuperaDatosConyuge(ByVal pnCodPers As Integer) As ComTipos.ClsPersona
        Dim C As ComTipos.ClsPersona
        C = New ComTipos.ClsPersona

        C.oDirec.oDatosDir.CargaTabla(, , "Cred_DevuelveDatosConyuge " & pnCodPers)
        C.Dependientes.CargaTabla(, , "Pers_RecuperaDependientes " & pnCodPers)
        C.oSucursales.CargaTabla(, , "Pers_RecuperaSucursales " & pnCodPers)
        C.Vivienda.CargaTabla(, , "Pers_RecuperaDatosPersona " & pnCodPers)
        RecuperaDatosConyuge = C
        C = Nothing
    End Function
    Public Function RecuperaControles() As ComTipos.ClsPersona
        Dim C As ComTipos.ClsPersona
        C = New ComTipos.ClsPersona

        C.EstadosCivil.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 100")
        C.RelaInst.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 300")
        C.Estados.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 200")
        C.EstadosPerJur.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2800")
        C.TiposDoc.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1400")
        C.TiposParent.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2900")
        C.TiposMagnEmp.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 3000")
        C.TiposPersonaJur.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 3001")
        C.GradosInt.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 3003")
        C.Profesiones.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 3004")
        C.oSiglas.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 4565")

        C.oPropiedad.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 2157")
        C.oPropietario.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 2158")
        C.oZona.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 2159")
        C.oBarrio.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 2160")
        C.oTipoViv.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 2161")
        C.oMaterial.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 2162")
        C.oEstadoViv.CargaTabla(, , " CredW_RecuperaCatalogoCodigos 2163")

        C.oDirec.oTipoDir1.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 800")
        C.oDirec.oTipoDir2.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 900")
        C.oDirec.oTipoDir3.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1000")
        C.oDirec.oTipoDir4.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1100")

        C.oDirecTrab.oTipoDir1.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 800")
        C.oDirecTrab.oTipoDir2.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 900")
        C.oDirecTrab.oTipoDir3.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1000")
        C.oDirecTrab.oTipoDir4.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1100")

        C.oNegocio.oTipoDir1.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 800")
        C.oNegocio.oTipoDir2.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 900")
        C.oNegocio.oTipoDir3.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1000")
        C.oNegocio.oTipoDir4.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1100")

        C.oNegocio.oLocal.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 4560")
        C.oDirec.oLocal.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 4560")

        RecuperaControles = C

        'C.Nacionalidad.DevuelveXMLLinQ()
        C = Nothing

    End Function

    Public Function RecuperaUbigeo(ByVal pcNomDpto As String, ByVal pcNomProv As String, ByVal pcNomDist As String, ByVal pcNomZona As String) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset

        C.CargaTabla(, , "Cred_DevuelveUbigeo '" & pcNomDpto & "', '" & pcNomProv & "', '" & pcNomDist & "', '" & pcNomZona & "'")

        RecuperaUbigeo = C
        C = Nothing

    End Function


    Private Function RecuperaInventario(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaInventario " & pnCodEval)


        RecuperaInventario = D
        D = Nothing
    End Function

    Private Function RecuperaCicloComprasVentas(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaCicloComprasVentas " & pnCodEval)
        RecuperaCicloComprasVentas = D

        D = Nothing
    End Function
    Private Function RecuperaDeudasFinanConyuge(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDeudasFinancConyuge " & pnCodEval & ",'" & psNumDoc & "'")

        RecuperaDeudasFinanConyuge = D
        D = Nothing
    End Function
    Private Function RecuperaDeudasFinan(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDeudasFinanc " & pnCodEval & ",'" & psNumDoc & "'")

        RecuperaDeudasFinan = D
        D = Nothing
    End Function

    Public Function RecuperaEvalVentasImpresion(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.ClsEval_Ventas

        Dim C As ComTipos.ClsEval_Ventas
        C = New ComTipos.ClsEval_Ventas
        C.oDataCicAnuNeg.CargaTabla(, , "Cred_RecuperaCicAnuNegImp " & pnCodEval)
        C.oDataVentas = RecuperaCicloComprasVentas(pnCodEval)
        C.oDeudaFinan = RecuperaDeudasFinan(pnCodEval, psNumDoc)
        C.oDeudaFinanCony = RecuperaDeudasFinanConyuge(pnCodEval, psNumDoc)
        C.oInventario = RecuperaInventario(pnCodEval)

        C.oDataInsumo = RecuperaHojasDeCostoInsumo1(pnCodEval)
        C.oDataInsumo2 = RecuperaHojasDeCostoInsumo2(pnCodEval)
        C.oDataInsumo3 = RecuperaHojasDeCostoInsumo3(pnCodEval)
        C.oDataInsumo4 = RecuperaHojasDeCostoInsumo4(pnCodEval)
        C.oDataMatPri = RecuperaHojasDeCostoMatPrim1(pnCodEval)
        C.oDataMatPri2 = RecuperaHojasDeCostoMatPrim2(pnCodEval)
        C.oDataMatPri3 = RecuperaHojasDeCostoMatPrim3(pnCodEval)
        C.oDataMatPri4 = RecuperaHojasDeCostoMatPrim4(pnCodEval)

        Call CargarProductosHojaCostos(pnCodEval, C.psProd1, C.psProd2, C.psProd3, C.psProd4)

        RecuperaEvalVentasImpresion = C
        C = Nothing
    End Function

    Public Sub CargarProductosHojaCostos(ByVal pnCodEval As Integer, ByRef psProd1 As String, _
                                        ByRef psProd2 As String, ByRef psProd3 As String, ByRef psProd4 As String)
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoProductos " & pnCodEval)

        psProd1 = D.ObtenerDato("cNomProd1")
        psProd2 = D.ObtenerDato("cNomProd2")
        psProd3 = D.ObtenerDato("cNomProd3")
        psProd4 = D.ObtenerDato("cNomProd4")
        D = Nothing
    End Sub

    Public Function RecuperaHojasDeCostoInsumo1(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo1 " & pnCodEval)
        RecuperaHojasDeCostoInsumo1 = D
        D = Nothing
    End Function
    Public Function RecuperaHojasDeCostoInsumo2(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo2 " & pnCodEval)
        RecuperaHojasDeCostoInsumo2 = D
        D = Nothing
    End Function
    Public Function RecuperaHojasDeCostoInsumo3(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo3 " & pnCodEval)
        RecuperaHojasDeCostoInsumo3 = D
        D = Nothing
    End Function
    Public Function RecuperaHojasDeCostoInsumo4(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo4 " & pnCodEval)
        RecuperaHojasDeCostoInsumo4 = D
        D = Nothing
    End Function

    Public Function RecuperaHojasDeCostoMatPrim1(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim1 " & pnCodEval)
        RecuperaHojasDeCostoMatPrim1 = D
        D = Nothing
    End Function
    Public Function RecuperaHojasDeCostoMatPrim2(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim2 " & pnCodEval)
        RecuperaHojasDeCostoMatPrim2 = D
        D = Nothing
    End Function
    Public Function RecuperaHojasDeCostoMatPrim3(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim3 " & pnCodEval)
        RecuperaHojasDeCostoMatPrim3 = D
        D = Nothing
    End Function
    Public Function RecuperaHojasDeCostoMatPrim4(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim4 " & pnCodEval)
        RecuperaHojasDeCostoMatPrim4 = D
        D = Nothing
    End Function


    Public Function RecuperaResumenVentas(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "DevuelveHojaCostoProd " & pnCodEval)
        RecuperaResumenVentas = D
        D = Nothing
    End Function
    Public Function RecuperaEvalVentas(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.ClsEval_Ventas
        Dim C As ComTipos.ClsEval_Ventas
        Dim D As ComTipos.clsDataset

        C = New ComTipos.ClsEval_Ventas
        'C.oCboTipoProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1501")
        'C.oCboMeses.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2500")
        'C.oCboPlazo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 500")
        'C.oCboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        'C.oDataCicAnuNeg.CargaTabla(, , "Cred_RecuperaCicAnuNeg " & pnCodEval)
        'C.oDataVentas = RecuperaCicloComprasVentas(pnCodEval)
        'C.oDeudaFinan = RecuperaDeudasFinan(pnCodEval, psNumDoc)
        'C.oDeudaFinanCony = RecuperaDeudasFinanConyuge(pnCodEval, psNumDoc)
        C.oInventario = RecuperaInventario(pnCodEval)
        C.oDataMatPri.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim1 " & pnCodEval)
        C.oDataMatPri2.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim2 " & pnCodEval)
        C.oDataMatPri3.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim3 " & pnCodEval)
        C.oDataMatPri4.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim4 " & pnCodEval)
        C.oDataMatPri5.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim5 " & pnCodEval)
        C.oDataMatPri6.CargaTabla(, , "Cred_RecuperaHojasDeCostoMatPrim6 " & pnCodEval)
        C.oDataInsumo.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo1 " & pnCodEval)
        C.oDataInsumo2.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo2 " & pnCodEval)
        C.oDataInsumo3.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo3 " & pnCodEval)
        C.oDataInsumo4.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo4 " & pnCodEval)
        C.oDataInsumo5.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo5 " & pnCodEval)
        C.oDataInsumo6.CargaTabla(, , "Cred_RecuperaHojasDeCostoInsumo6 " & pnCodEval)
        C.oDataResumen.CargaTabla(, , "DevuelveHojaCostoProd " & pnCodEval)

        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojasDeCostoProductos " & pnCodEval)

        C.psProd1 = D.ObtenerDato("cNomProd1")
        C.psProd2 = D.ObtenerDato("cNomProd2")
        C.psProd3 = D.ObtenerDato("cNomProd3")
        C.psProd4 = D.ObtenerDato("cNomProd4")
        C.psProd5 = D.ObtenerDato("cNomProd5")
        C.psProd6 = D.ObtenerDato("cNomProd6")


        RecuperaEvalVentas = C
        C = Nothing
    End Function
    Public Function RecuperaEvalDeudasFinancieras(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.ClsEval_Ventas
        Dim C As ComTipos.ClsEval_Ventas

        C = New ComTipos.ClsEval_Ventas
        C.oCboTipoProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1501")
        C.oCboPlazo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 500")
        C.oCboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.oDeudaFinan = RecuperaDeudasFinan(pnCodEval, psNumDoc)
        C.oDeudaFinanCony = RecuperaDeudasFinanConyuge(pnCodEval, psNumDoc)


        RecuperaEvalDeudasFinancieras = C
        C = Nothing
    End Function

    Public Sub RegistraEvalVentas(ByVal pnCodEval As Integer, ByVal D As ComTipos.ClsEval_Ventas, ByVal pnCant1 As Integer, ByVal pnCant2 As Integer, ByVal pnCant3 As Integer, ByVal pnCant4 As Integer, ByVal pnCant5 As Integer, ByVal pnCant6 As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RegistraEvalVentas", objCon.DameConexion)
        'oComm = New SqlClient.SqlCommand("Cred_RegistraEvalVentas")
        'oComm.Connection = objCon.DameConexion
        'objCon.IniciaTransaccion()
        'oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@CicloAnuNeg", SqlDbType.Xml).Value = D.oDataCicAnuNeg.DevuelveXML
        oComm.Parameters.Add("@Ventas", SqlDbType.Xml).Value = D.oDataVentas.DevuelveXML
        oComm.Parameters.Add("@Inventario", SqlDbType.Xml).Value = D.oInventario.DevuelveXML

        'ARCV 11-05-2011
        oComm.ExecuteNonQuery()
        'objCon.EjecutaTransaccion(oComm)
        'objCon.FinalizaTransaccion()
        '-----------

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing

        RegistroHojaCostos(pnCodEval, D.psProd1, D.psProd2, D.psProd3, D.psProd4, D.psProd5, D.psProd6, _
            D.oDataMatPri, D.oDataMatPri2, D.oDataMatPri3, D.oDataMatPri4, D.oDataMatPri5, D.oDataMatPri6, _
            D.oDataInsumo, D.oDataInsumo2, D.oDataInsumo3, D.oDataInsumo4, D.oDataInsumo5, D.oDataInsumo6, pnCant1, pnCant2, pnCant3, pnCant4, pnCant5, pnCant6)

    End Sub

    Public Sub RegistraEvalDeudasFinancieras(ByVal pnCodEval As Integer, ByVal D As ComTipos.ClsEval_Ventas)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Cred_RegistraEvalDeudasFinancieras", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_RegistraEvalDeudasFinancieras")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@DeudaFinan", SqlDbType.Xml).Value = D.oDeudaFinan.DevuelveXML
        oComm.Parameters.Add("@DeudaFinanCony", SqlDbType.Xml).Value = D.oDeudaFinanCony.DevuelveXML

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing

    End Sub
    Public Sub GrabarPersona(ByVal oTipo As ComTipos.clsCredPersonaNat)
        Dim oCon As ComConecta.clsConecta
        Dim i As Integer
        'Dim nCodPers As Integer

        'oCon = New ComConecta.clsConecta
        'Dim oComm As New SqlCommand("Cred_PersonaNat_Insertar", oCon.DameConexion)
        'oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = oTipo.nCodAge
        'oComm.Parameters.Add("@nCodPers", SqlDbType.Int)
        'oComm.Parameters.Item("@nCodPers").Direction = ParameterDirection.Output
        'oComm.Parameters.Add("@nDirTipo1", SqlDbType.Int).Value = oTipo.nDirTipo1
        'oComm.Parameters.Add("@nDirTipo2", SqlDbType.Int).Value = oTipo.nDirTipo2
        'oComm.Parameters.Add("@nDirTipo3", SqlDbType.Int).Value = oTipo.nDirTipo3
        'oComm.Parameters.Add("@nDirTipo4", SqlDbType.Int).Value = oTipo.nDirTipo4
        'oComm.Parameters.Add("@cDirValor1", SqlDbType.VarChar).Value = oTipo.cDirValor1
        'oComm.Parameters.Add("@cDirValor2", SqlDbType.VarChar).Value = oTipo.cDirValor2
        'oComm.Parameters.Add("@cDirValor3", SqlDbType.VarChar).Value = oTipo.cDirValor3
        'oComm.Parameters.Add("@cDirValor4", SqlDbType.VarChar).Value = oTipo.cDirValor4
        'oComm.Parameters.Add("@cDirReferencia", SqlDbType.VarChar).Value = oTipo.cDirReferencia
        'oComm.Parameters.Add("@cZona", SqlDbType.VarChar).Value = oTipo.cZona
        'oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar).Value = oTipo.cTelefono
        'oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = oTipo.nEstado
        'oComm.Parameters.Add("@nTipoPersona", SqlDbType.Int).Value = oTipo.nTipoPersona
        'oComm.Parameters.Add("@cCIIU", SqlDbType.VarChar).Value = oTipo.cCIIU
        'oComm.Parameters.Add("@cApePat", SqlDbType.VarChar).Value = oTipo.cApePat
        'oComm.Parameters.Add("@cApeMat", SqlDbType.VarChar).Value = oTipo.cApeMat
        'oComm.Parameters.Add("@cApeCas", SqlDbType.VarChar).Value = oTipo.cApeCas
        'oComm.Parameters.Add("@cNombres", SqlDbType.VarChar).Value = oTipo.cNombres
        'oComm.Parameters.Add("@cSexo", SqlDbType.VarChar).Value = oTipo.cSexo
        'oComm.Parameters.Add("@dFecNac", SqlDbType.DateTime).Value = oTipo.dFecNac
        'oComm.Parameters.Add("@cMovil", SqlDbType.VarChar).Value = oTipo.cMovil
        'oComm.Parameters.Add("@cDNI", SqlDbType.VarChar).Value = oTipo.cDNI
        'oComm.Parameters.Add("@nEstadoCivil", SqlDbType.Int).Value = oTipo.nEstadoCivil
        'oComm.Parameters.Add("@nNacionalidad", SqlDbType.Int).Value = oTipo.nNacionalidad
        'oComm.Parameters.Add("@nRelacIns", SqlDbType.Int).Value = oTipo.nNacionalidad
        'oComm.Parameters.Add("@cRpm_RPC", SqlDbType.VarChar).Value = oTipo.cRpm_RPC
        'oComm.Parameters.Add("@cMail", SqlDbType.VarChar).Value = oTipo.cMail
        'oComm.Parameters.Add("@bParentesco", SqlDbType.Bit).Value = oTipo.bParentesco
        'oComm.Parameters.Add("@bImpactoSocial", SqlDbType.Bit).Value = oTipo.bImpactoSocial
        'oComm.ExecuteNonQuery()

        'nCodPers = oComm.Parameters.Item("@nCodPers").Value

        'For i = 1 To UBound(oTipo.Documentos)
        '    oComm = New SqlCommand("Cred_PersonaDoc_Insertar", oCon.ConexionActual)
        '    oComm.CommandType = CommandType.StoredProcedure
        '    oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = oTipo.nCodAge
        '    oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        '    oComm.Parameters.Add("@nTipoDoc", SqlDbType.Int).Value = CInt(oTipo.Documentos(i).sCodigo)
        '    oComm.Parameters.Add("@cValor", SqlDbType.VarChar).Value = oTipo.Documentos(i).sDescripcion

        '    oComm.ExecuteNonQuery()
        'Next

        'oComm.Parameters.Add("@Table", SqlDbType.Xml)

        'For i = 1 To UBound(oTipo.Parentesco)
        '    oComm = New SqlCommand("Cred_PersonaRelac_Insertar", oCon.ConexionActual)
        '    oComm.CommandType = CommandType.StoredProcedure
        '    oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = oTipo.nCodAge
        '    oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        '    oComm.Parameters.Add("@nTipoRela", SqlDbType.Int).Value = oTipo.Parentesco(i).nTipoRelac
        '    oComm.Parameters.Add("@nCodAgeRela", SqlDbType.Int).Value = oTipo.Parentesco(i).nCodAgeRela
        '    oComm.Parameters.Add("@nCodPersRela", SqlDbType.Int).Value = oTipo.Parentesco(i).nCodPersRela
        '    oComm.ExecuteNonQuery()
        'Next

        'oCon.CierraConexion()
        'oComm = Nothing

    End Sub

    Public Function BuscarPersona(ByVal nTipoBusqueda As Integer, _
                                ByVal sCriterioBusqueda As String, _
                                ByVal nTipoPersona As Integer) As ComTipos.clsBuscaPersona

        Dim oTipo As New ComTipos.clsBuscaPersona
        Dim oComm As SqlCommand
        Dim oCon As ComConecta.clsConecta
        Dim da As SqlDataReader

        oCon = New ComConecta.clsConecta

        oComm = New SqlCommand("Cred_Persona_Buscar", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nTipoBusq", SqlDbType.Int).Value = nTipoBusqueda
        oComm.Parameters.Add("@cCriterioBusq", SqlDbType.VarChar).Value = sCriterioBusqueda
        oComm.Parameters.Add("@nTipoPer", SqlDbType.Int).Value = nTipoPersona
        da = oComm.ExecuteReader(CommandBehavior.SingleResult)

        ReDim oTipo.Resultado(0)
        While da.Read
            ReDim Preserve oTipo.Resultado(UBound(oTipo.Resultado) + 1)
            oTipo.Resultado(UBound(oTipo.Resultado)).nCodAge = da.GetString(0)
            oTipo.Resultado(UBound(oTipo.Resultado)).nCodPers = da.GetString(1)
            oTipo.Resultado(UBound(oTipo.Resultado)).sDescripcion = da.GetString(2)
            oTipo.Resultado(UBound(oTipo.Resultado)).sDocumento = da.GetString(3)
        End While

        oCon.CierraConexion()
        oComm = Nothing

        BuscarPersona = oTipo
    End Function

    'Public Sub CargarGrupoCIIU(ByVal nGrupo As Integer, _
    '                           ByVal sBuscar As String, _
    '                           ByRef P As ComTipos.clsCargaPersona)

    '    Dim oComm As SqlCommand
    '    Dim oCon As ComConecta.clsConecta
    '    Dim da As SqlDataReader
    '    oCon = New ComConecta.clsConecta

    '    oComm = New SqlCommand("Cred_CIIU_Recupera20", oCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    oComm.Parameters.Add("@nGrupo", SqlDbType.Int).Value = nGrupo
    '    oComm.Parameters.Add("@cBuscar", SqlDbType.VarChar).Value = sBuscar
    '    da = oComm.ExecuteReader(CommandBehavior.SingleResult)

    '    ReDim P.ActividadEcon(0)
    '    While da.Read
    '        ReDim Preserve P.ActividadEcon(UBound(P.ActividadEcon) + 1)
    '        P.ActividadEcon(UBound(P.ActividadEcon)).sCodigo = da.GetString(0)
    '        P.ActividadEcon(UBound(P.ActividadEcon)).sDescripcion = da.GetString(1)
    '    End While

    '    oCon.CierraConexion()
    '    oComm = Nothing

    'End Sub

    'Public Sub CargarGrupoPais(ByVal nGrupo As Integer, _
    '                       ByVal sBuscar As String, _
    '                       ByRef P As ComTipos.clsCargaPersona)

    '    Dim oComm As SqlCommand
    '    Dim oCon As ComConecta.clsConecta
    '    Dim da As SqlDataReader

    '    oCon = New ComConecta.clsConecta
    '    oComm = New SqlCommand("Cred_Paises_Recupera20", oCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    oComm.Parameters.Add("@nGrupo", SqlDbType.Int).Value = nGrupo
    '    oComm.Parameters.Add("@cBuscar", SqlDbType.VarChar).Value = sBuscar
    '    da = oComm.ExecuteReader(CommandBehavior.SingleResult)

    '    ReDim P.Nacionalidad(0)
    '    While da.Read
    '        ReDim Preserve P.Nacionalidad(UBound(P.Nacionalidad) + 1)
    '        P.Nacionalidad(UBound(P.Nacionalidad)).sCodigo = da.GetString(0)
    '        P.Nacionalidad(UBound(P.Nacionalidad)).sDescripcion = da.GetString(1)
    '    End While

    '    oCon.CierraConexion()
    '    oComm = Nothing

    'End Sub

    Public Sub CargarDatosFormPersona(ByRef P As ComTipos.clsCargaPersona)

        'Dim oComm As SqlCommand
        'Dim oCon As ComConecta.clsConecta
        'Dim da As SqlDataReader

        oCons = New ComConstante.clsConstante
        P.Estado = oCons.RecuperaConstantes(200)  'oCons.RecuperaConstantesADO(200)
        P.RelInstitucional = oCons.RecuperaConstantes(300)
        P.EstadoCivil = oCons.RecuperaConstantes(100)
        'oCons.RecuperaConstantes(100, P.EstadoCivil)
        'oCons.RecuperaConstantes(200, P.Estado)
        'oCons.RecuperaConstantes(300, P.RelInstitucional)
        oCons = Nothing

        'oCon = New ComConecta.clsConecta

        'oComm = New SqlCommand("Cred_Paises_Recupera20", oCon.DameConexion)
        'oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@nGrupo", SqlDbType.Int).Value = 1
        'oComm.Parameters.Add("@cBuscar", SqlDbType.VarChar).Value = ""
        'da = oComm.ExecuteReader(CommandBehavior.SingleResult)

        'ReDim P.Nacionalidad(0)
        'While da.Read
        '    ReDim Preserve P.Nacionalidad(UBound(P.Nacionalidad) + 1)
        '    P.Nacionalidad(UBound(P.Nacionalidad)).sCodigo = da.GetString(0)
        '    P.Nacionalidad(UBound(P.Nacionalidad)).sDescripcion = da.GetString(1)
        'End While

        'oComm = New SqlCommand("Cred_CIIU_Recupera20", oCon.DameConexion)
        'oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@nGrupo", SqlDbType.Int).Value = 1
        'oComm.Parameters.Add("@cBuscar", SqlDbType.VarChar).Value = ""
        'da = oComm.ExecuteReader(CommandBehavior.SingleResult)

        'ReDim P.ActividadEcon(0)
        'While da.Read
        '    ReDim Preserve P.ActividadEcon(UBound(P.ActividadEcon) + 1)
        '    P.ActividadEcon(UBound(P.ActividadEcon)).sCodigo = da.GetString(0)
        '    P.ActividadEcon(UBound(P.ActividadEcon)).sDescripcion = da.GetString(1)
        'End While
        'oCon.CierraConexion()
        'oComm = Nothing
    End Sub

    'Public Function CargarPersonaDirForm() As ComTipos.clsCredPersonaDir
    '    Dim D As New ComTipos.clsCredPersonaDir

    '    Dim oComm As SqlCommand
    '    Dim oCon As ComConecta.clsConecta
    '    Dim da As SqlDataReader
    '    oCons = New ComConstante.clsConstante

    '    oCons.RecuperaConstantes(800, D.TipoDir1)
    '    oCons.RecuperaConstantes(900, D.TipoDir2)
    '    oCons.RecuperaConstantes(1000, D.TipoDir3)
    '    oCons.RecuperaConstantes(1100, D.TipoDir4)
    '    oCons = Nothing

    '    oCon = New ComConecta.clsConecta
    '    oComm = New SqlCommand("Cred_Departamento_Recupera", oCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    da = oComm.ExecuteReader(CommandBehavior.SingleResult)

    '    ReDim D.Departamentos(0)
    '    While da.Read
    '        ReDim Preserve D.Departamentos(UBound(D.Departamentos) + 1)
    '        D.Departamentos(UBound(D.Departamentos)).sCodigo = da.GetString(0)
    '        D.Departamentos(UBound(D.Departamentos)).sDescripcion = da.GetString(1)
    '    End While

    '    oCon.CierraConexion()
    '    oComm = Nothing

    '    Return D

    'End Function

    'Public Function CargarZonas(ByVal nGrupo As Integer, _
    '                            ByVal cBuscar As String, _
    '                            ByVal nTipoBusq As Integer, _
    '                            ByVal cPadre As String) As ComTipos.clsTipoCombo()
    '    Dim Z() As ComTipos.clsTipoCombo

    '    Dim oComm As SqlCommand
    '    Dim oCon As ComConecta.clsConecta
    '    Dim da As SqlDataReader

    '    oCon = New ComConecta.clsConecta
    '    oComm = New SqlCommand("Cred_Zonas_Recupera20", oCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    oComm.Parameters.Add("@nGrupo", SqlDbType.Int).Value = nGrupo
    '    oComm.Parameters.Add("@cBuscar", SqlDbType.VarChar).Value = cBuscar
    '    oComm.Parameters.Add("@nTipoBusqueda", SqlDbType.Int).Value = nTipoBusq
    '    oComm.Parameters.Add("@cPadre", SqlDbType.VarChar).Value = cPadre
    '    da = oComm.ExecuteReader(CommandBehavior.SingleResult)

    '    ReDim Z(0)
    '    While da.Read
    '        ReDim Preserve Z(UBound(Z) + 1)
    '        Z(UBound(Z)).sCodigo = da.GetString(0)
    '        Z(UBound(Z)).sDescripcion = da.GetString(1)
    '    End While

    '    oCon.CierraConexion()
    '    oComm = Nothing

    '    Return Z

    'End Function

    'Public Function CargarPersonaDocForm() As ComTipos.clsCredPersonaDoc
    '    Dim D As New ComTipos.clsCredPersonaDoc
    '    oCons = New ComConstante.clsConstante

    '    oCons.RecuperaConstantes(1400, D.Documentos)
    '    oCons = Nothing

    '    Return D

    'End Function

    Public Function CredEvalFlujoCajaCRecuperaIngresos(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredEvalFCRecupIngr " & pnCodEval)
        CredEvalFlujoCajaCRecuperaIngresos = C
        C = Nothing
    End Function

    Public Function CredEvalFlujoCajaCRecuperaEgresos(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredEvalFCRecupEgr " & pnCodEval)
        CredEvalFlujoCajaCRecuperaEgresos = C
        C = Nothing
    End Function

    Public Function RecuperaEvalUniFamImpresion(ByVal pnCodEval As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaEvalUniFamImpresion " & pnCodEval)
        RecuperaEvalUniFamImpresion = D
        D = Nothing
    End Function
    Public Function RecuperaEvalUniFam(ByVal pnCodEval As Integer, ByVal pnCodPers As Integer) As ComTipos.ClsEval_UniFam
        Dim D As ComTipos.ClsEval_UniFam
        D = New ComTipos.ClsEval_UniFam
        D.oDisCli.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2200")
        D.oExpNeg.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2201")
        D.oRieZonNeg.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2202")
        D.oRefCom.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2203")
        D.oComp.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2204")
        D.oCtasCobAtr.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2205")
        D.oGraInst.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2206")
        D.oEdad.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2207")
        D.oRieZonDom.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2208")
        D.oSolFam.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2209")
        D.oAntCred.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2210")
        D.oDatos.CargaTabla(, , " Cred_RecuperaEvalUniFam " & pnCodEval.ToString)
        D.oDependientes.CargaTabla(, , " Cred_RecuperaDependientes " & pnCodPers)

        RecuperaEvalUniFam = D
        D = Nothing

    End Function

    Public Sub RegistraEvalUniFam(ByVal D As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        'ARCV 11-05-2011
        objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Cred_RegistraEvalUniFam", objCon.DameConexion)
        oComm = New SqlClient.SqlCommand("Cred_RegistraEvalUniFam")
        oComm.Connection = objCon.DameConexion
        objCon.IniciaTransaccion()
        oComm.Transaction = objCon.TransaccionActual
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = D.ObtenerDato("nCodEval")
        oComm.Parameters.Add("@nDisCli", SqlDbType.Int).Value = D.ObtenerDato("nDisCli")
        oComm.Parameters.Add("@nExpNeg", SqlDbType.Int).Value = D.ObtenerDato("nExpNeg")
        oComm.Parameters.Add("@nRieZonNeg", SqlDbType.Int).Value = D.ObtenerDato("nRieZonNeg")
        oComm.Parameters.Add("@nRefCom", SqlDbType.Int).Value = D.ObtenerDato("nRefCom")
        oComm.Parameters.Add("@nComp", SqlDbType.Int).Value = D.ObtenerDato("nComp")
        oComm.Parameters.Add("@nCueCob", SqlDbType.Int).Value = D.ObtenerDato("nCueCob")
        oComm.Parameters.Add("@nGradInst", SqlDbType.Int).Value = D.ObtenerDato("nGradInst")
        oComm.Parameters.Add("@nEdadTit", SqlDbType.Int).Value = D.ObtenerDato("nEdadTit")
        oComm.Parameters.Add("@nRieZonDom", SqlDbType.Int).Value = D.ObtenerDato("nRieZonDom")
        oComm.Parameters.Add("@nSolFam", SqlDbType.Int).Value = D.ObtenerDato("nSolFam")
        oComm.Parameters.Add("@nAntCre", SqlDbType.Int).Value = D.ObtenerDato("nAntCre")
        oComm.Parameters.Add("@nIngAdi", SqlDbType.Money).Value = D.ObtenerDato("nIngAdi")
        oComm.Parameters.Add("@nOtrosIng", SqlDbType.Money).Value = D.ObtenerDato("nOtrosIng")
        oComm.Parameters.Add("@nAlimentac", SqlDbType.Money).Value = D.ObtenerDato("nAlimentac")
        oComm.Parameters.Add("@nAlqCasa", SqlDbType.Money).Value = D.ObtenerDato("nAlqCasa")
        oComm.Parameters.Add("@nVesCal", SqlDbType.Money).Value = D.ObtenerDato("nVesCal")
        oComm.Parameters.Add("@nEducacion", SqlDbType.Money).Value = D.ObtenerDato("nEducacion")
        oComm.Parameters.Add("@nSalud", SqlDbType.Money).Value = D.ObtenerDato("nSalud")
        oComm.Parameters.Add("@nTransporte", SqlDbType.Money).Value = D.ObtenerDato("nTransporte")
        oComm.Parameters.Add("@nServicios", SqlDbType.Money).Value = D.ObtenerDato("nServicios")
        oComm.Parameters.Add("@nOtros", SqlDbType.Money).Value = D.ObtenerDato("nOtros")
        oComm.Parameters.Add("@nCuotaMens", SqlDbType.Money).Value = D.ObtenerDato("nCuotaMens")
        oComm.Parameters.Add("@nJuntas", SqlDbType.Money).Value = D.ObtenerDato("nJuntas")
        oComm.Parameters.Add("@nNumDep", SqlDbType.Int).Value = D.ObtenerDato("nNumDep")
        oComm.Parameters.Add("@cEdades", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cEdades")
        oComm.Parameters.Add("@cComenOtrIng", SqlDbType.VarChar, 1000).Value = D.ObtenerDato("cComenOtrIng")
        oComm.Parameters.Add("@nIngCony", SqlDbType.Money).Value = D.ObtenerDato("nIngCony")
        oComm.Parameters.Add("@nTitIng", SqlDbType.Money).Value = D.ObtenerDato("nTitIng")
        oComm.Parameters.Add("@nConyIng", SqlDbType.Money).Value = D.ObtenerDato("nConyIng")
        oComm.Parameters.Add("@nAlqIng", SqlDbType.Money).Value = D.ObtenerDato("nAlqIng")
        oComm.Parameters.Add("@nRemesasIng", SqlDbType.Money).Value = D.ObtenerDato("nRemesasIng")

        '02/08/2012 eduenas
        oComm.Parameters.Add("@nServiciosExtras", SqlDbType.Money).Value = D.ObtenerDato("nServiciosExtras")
        oComm.Parameters.Add("@nOtrosImp", SqlDbType.Money).Value = D.ObtenerDato("nOtrosImp")
        oComm.Parameters.Add("@nPagosOtros", SqlDbType.Money).Value = D.ObtenerDato("nPagosOtros")

        'ARCV 11-05-2011
        'oComm.ExecuteNonQuery()
        objCon.EjecutaTransaccion(oComm)
        objCon.FinalizaTransaccion()
        '-----------

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing


    End Sub

    Public Sub RegistraFlujoCaja(ByVal pnCodEval As Integer, ByVal oDatos As ComTipos.ClsEval_FlujoCaja)
        Dim oCon As ComConecta.clsConecta

        oCon = New ComConecta.clsConecta
        Dim oComm As New SqlCommand("Cred_RegistraFlujoCaja", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@pGridIngr", SqlDbType.Xml).Value = oDatos.oGridIngresos.DevuelveXML
        oComm.Parameters.Add("@pGridEgr", SqlDbType.Xml).Value = oDatos.oGridEgresos.DevuelveXML
        oComm.Parameters.Add("@pGridIncPor", SqlDbType.Xml).Value = oDatos.oGridIncVentas.DevuelveXML

        oComm.ExecuteNonQuery()
        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

    End Sub

    Public Function Relac_ValidaEdad(ByVal nCondCred As Integer, _
                                ByVal oDatos As ComTipos.clsDataset) As ComTipos.clsDataset
        Dim oCon As ComConecta.clsConecta
        oCon = New ComConecta.clsConecta
        Dim oReader As SqlDataReader
        Dim nNumReg As Integer

        Dim oResul As ComTipos.clsDataset
        oResul = New ComTipos.clsDataset

        Dim oComm As New SqlCommand("Pers_Relac_ValidaEdad", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@oDatos", SqlDbType.Xml).Value = oDatos.DevuelveXML
        oComm.Parameters.Add("@nCondCred", SqlDbType.Int).Value = nCondCred

        oComm.Parameters.Add("@nNumRegs", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        nNumReg = CInt(oComm.Parameters.Item("@nNumRegs").Value)

        oReader = oComm.ExecuteReader(CommandBehavior.SingleResult)

        Call oResul.LlenarConSqlDataReader(oReader, nNumReg)
        oReader.Close()
        oReader = Nothing

        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

        Return oResul

    End Function

    Public Function RecuperaZonasRegistrales_11356() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Pers_RecuperaZonasRegistrales_11356 ")
        RecuperaZonasRegistrales_11356 = D
        D = Nothing
    End Function

    Public Function RegistraPersona_11356(ByVal pnCodPers As Integer, ByVal pnCodAge As Integer, ByVal poData As ComTipos.clsDataset, ByVal pDocs As ComTipos.clsDataset, ByVal pParen As ComTipos.clsDataset, ByVal pDepen As ComTipos.clsDataset, ByVal pSucursal As ComTipos.clsDataset, ByVal pRepresentantes As ComTipos.clsDataset, ByVal pAccionistas As ComTipos.clsDataset) As Integer

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Pers_RegistraPersona_11356", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pData", SqlDbType.Xml).Value = poData.DevuelveXML
        oComm.Parameters.Add("@pDocs", SqlDbType.Xml).Value = pDocs.DevuelveXML
        oComm.Parameters.Add("@pParen", SqlDbType.Xml).Value = pParen.DevuelveXML
        oComm.Parameters.Add("@pDepen", SqlDbType.Xml).Value = pDepen.DevuelveXML
        oComm.Parameters.Add("@pSucursal", SqlDbType.Xml).Value = pSucursal.DevuelveXML
        oComm.Parameters.Add("@pRepresentantes", SqlDbType.Xml).Value = pRepresentantes.DevuelveXML
        oComm.Parameters.Add("@pAccionistas", SqlDbType.Xml).Value = pAccionistas.DevuelveXML
        oComm.Parameters.Add("@pnCodPersNew", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        RegistraPersona_11356 = oComm.Parameters("@pnCodPersNew").Value

        oComm = Nothing
        objCon.CierraConexion()
        objCon = Nothing

    End Function

    Public Function Relac_ValidaEdadGrupo(ByVal nCondCred As Integer, _
                            ByVal oDatos As ComTipos.clsDataset) As ComTipos.clsDataset
        Dim oCon As ComConecta.clsConecta
        oCon = New ComConecta.clsConecta
        Dim oReader As SqlDataReader
        Dim nNumReg As Integer

        Dim oResul As ComTipos.clsDataset
        oResul = New ComTipos.clsDataset

        Dim oComm As New SqlCommand("Pers_Relac_ValidaEdadGrupo", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@oDatos", SqlDbType.Xml).Value = oDatos.DevuelveXML
        oComm.Parameters.Add("@nCondCred", SqlDbType.Int).Value = nCondCred

        oComm.Parameters.Add("@nNumRegs", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        nNumReg = CInt(oComm.Parameters.Item("@nNumRegs").Value)

        oReader = oComm.ExecuteReader(CommandBehavior.SingleResult)

        Call oResul.LlenarConSqlDataReader(oReader, nNumReg)
        oReader.Close()
        oReader = Nothing

        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

        Return oResul

    End Function

    Public Function BuscarPersonaclsDataset(ByVal nTipoBusqueda As Integer, _
                            ByVal sCriterioBusqueda As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_Persona_Buscar " & nTipoBusqueda & ",'" & sCriterioBusqueda & "'")
        BuscarPersonaclsDataset = D
        D = Nothing
    End Function

    Public Function DevuelveEdad(ByVal pdFecNac As Date, ByVal pdFecActual As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Log_DevuelveEdad '" & pdFecNac & "', '" & pdFecActual & "' ")
        DevuelveEdad = D
        D = Nothing
    End Function

    Public Function RecuperaBusquedaRCCTotalNombres(ByVal pnTipoBus As Integer, ByVal pnTipo As Integer, ByVal pcApePat As String, _
    ByVal pcApeMat As String, ByVal pcNombres As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_BusquedaPersonaNombresRCCTotal " & pnTipoBus & ", " & pnTipo & ", '" & pcApePat & "', '" & pcApeMat & "', '" & pcNombres & "' ")
        RecuperaBusquedaRCCTotalNombres = D
        D = Nothing
    End Function

    Public Function RecuperaBusquedaRCCTotalDocumento(ByVal pnTipoBus As Integer, ByVal pnTipo As Integer, ByVal pcDocumento As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_BusquedaPersonaNombresRCCTotal " & pnTipoBus & ", " & pnTipo & ", '', '', '', '" & pcDocumento & "' ")
        RecuperaBusquedaRCCTotalDocumento = D
        D = Nothing
    End Function

    Public Function RecuperaDataDeudasFinancieras(ByVal pnCodEval As Integer) As ComTipos.ClsEval_Ventas
        Dim C As ComTipos.ClsEval_Ventas

        C = New ComTipos.ClsEval_Ventas
        C.oCboTipoProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1501")
        C.oCboPlazo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 500")
        C.oCboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.oDeudaFinan.CargaTabla(, , "Pers_RecuperaDeudasFinanciera " & pnCodEval & ",1")
        C.oDeudaFinanCony.CargaTabla(, , "Pers_RecuperaDeudasFinanciera " & pnCodEval & ",2")
        'C.oDeudaFinanCony = RecuperaDeudasFinanConyuge(pnCodEval, psNumDoc)
        RecuperaDataDeudasFinancieras = C
        C = Nothing
    End Function

    Public Sub RegistraDeudaFinanciera(ByVal pnCodEval As Integer, _
    ByVal pnTipoPersona As Integer, ByVal pnSaldo As Double, ByVal pdFecVenc As Date, ByVal pcCuotaPag As String, _
    ByVal pnCuotaMens As Double, ByVal pcFinanc As String, ByVal pcClasificacion As String, ByVal pnPrestamo As Double, _
    ByVal pnDestino As Integer, ByVal pnPlazo As Integer, ByVal pcMoneda As String, ByVal pnCuotaFam As Double, _
    ByVal pnCuotaNego As Double, ByVal pnCapAFecha As Double, ByVal pnTipoProd As Integer, ByVal pnCapFecLP As Double)

        Dim oCon As ComConecta.clsConecta

        oCon = New ComConecta.clsConecta
        Dim oComm As New SqlCommand("Pers_RegistraDeudaFinanc", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@nTipoPersona", SqlDbType.Int).Value = pnTipoPersona
        oComm.Parameters.Add("@nSaldo", SqlDbType.Money).Value = pnSaldo
        oComm.Parameters.Add("@dFecVenc", SqlDbType.Date).Value = pdFecVenc
        oComm.Parameters.Add("@cCuotaPag", SqlDbType.VarChar, 8).Value = pcCuotaPag
        oComm.Parameters.Add("@nCuotaMens", SqlDbType.Money).Value = pnCuotaMens
        oComm.Parameters.Add("@cFinanc", SqlDbType.VarChar, 100).Value = pcFinanc
        oComm.Parameters.Add("@cClasific", SqlDbType.VarChar, 30).Value = pcClasificacion
        oComm.Parameters.Add("@nPrestamos", SqlDbType.Money).Value = pnPrestamo
        oComm.Parameters.Add("@nDestino", SqlDbType.Int).Value = pnDestino
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int).Value = pnPlazo
        oComm.Parameters.Add("@cMoneda", SqlDbType.VarChar, 20).Value = pcMoneda
        oComm.Parameters.Add("@nCuotaFam", SqlDbType.Money).Value = pnCuotaFam
        oComm.Parameters.Add("@nCuotaNeg", SqlDbType.Money).Value = pnCuotaNego
        oComm.Parameters.Add("@nCapAFec", SqlDbType.Money).Value = pnCapAFecha
        oComm.Parameters.Add("@nTipoProd", SqlDbType.Int).Value = pnTipoProd
        oComm.Parameters.Add("@nCapAFechaLP", SqlDbType.Money).Value = pnCapFecLP

        oComm.ExecuteNonQuery()
        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

    End Sub

    Public Sub EliminaDeudaFinanciera(ByVal pnCodEval As Integer, ByVal pnTipoPersona As Integer, ByVal pdFecVenc As Date, _
  ByVal pcFinanc As String, ByVal pnPrestamo As Double)

        Dim oCon As ComConecta.clsConecta

        oCon = New ComConecta.clsConecta
        Dim oComm As New SqlCommand("Pers_EliminaDeudaFinanc", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@nTipoPersona", SqlDbType.Int).Value = pnTipoPersona
        oComm.Parameters.Add("@dFecVenc", SqlDbType.Date).Value = pdFecVenc
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = pnPrestamo
        oComm.Parameters.Add("@cFinanc", SqlDbType.VarChar, 100).Value = pcFinanc

        oComm.ExecuteNonQuery()
        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

    End Sub

    Public Function ListaActivosNoCorrientes(ByVal pnCodEval As Integer, ByVal pntipo As Integer) As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Pers_ListaActivos " & pnCodEval & "," & pntipo)
        ListaActivosNoCorrientes = C
        C = Nothing
    End Function
    Public Function ListaEstadoActNoCorr() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 71")
        ListaEstadoActNoCorr = C
        C = Nothing
    End Function

    Public Function ListaTiposActNoCorr() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 72")
        ListaTiposActNoCorr = C
        C = Nothing
    End Function

    Public Sub RegistraActivo(ByVal pnCodEval As Integer, ByVal pnTipo As Integer, ByVal pcDescripcion As String, _
 ByVal pnAnio As Integer, ByVal pnValor As Double, ByVal pnEstado As Integer)

        Dim oCon As ComConecta.clsConecta

        oCon = New ComConecta.clsConecta
        Dim oComm As New SqlCommand("PersRegistraActivo", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = pnTipo
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Value = pcDescripcion
        oComm.Parameters.Add("@nAnio", SqlDbType.Int).Value = pnAnio
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = pnEstado
        oComm.Parameters.Add("@nValor", SqlDbType.Money).Value = pnValor
        oComm.ExecuteNonQuery()
        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

    End Sub


    Public Sub ActualizaActivo(ByVal pnCodEval As Integer, ByVal pnTipo As Integer, ByVal pnItem As Integer, ByVal pcDescripcion As String, _
    ByVal pnAnio As Integer, ByVal pnValor As Double, ByVal pnEstado As Integer)

        Dim oCon As ComConecta.clsConecta
        oCon = New ComConecta.clsConecta
        Dim oComm As New SqlCommand("PersActualizaActivo", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = pnTipo
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = pnItem
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Value = pcDescripcion
        oComm.Parameters.Add("@nAnio", SqlDbType.Int).Value = pnAnio
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = pnEstado
        oComm.Parameters.Add("@nValor", SqlDbType.Money).Value = pnValor
        oComm.ExecuteNonQuery()
        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

    End Sub

    Public Sub EliminaActivo(ByVal pnCodEval As Integer, ByVal pnTipo As Integer, ByVal pnItem As Integer)
        Dim oCon As ComConecta.clsConecta
        oCon = New ComConecta.clsConecta
        Dim oComm As New SqlCommand("PersEliminaActivo", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = pnTipo
        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = pnItem

        oComm.ExecuteNonQuery()
        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

    End Sub


    Public Function RecuperaEvalDeudasFinancieras_NewEval(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.ClsEval_Ventas
        Dim C As ComTipos.ClsEval_Ventas

        C = New ComTipos.ClsEval_Ventas
        C.oCboTipoProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 1501")
        C.oCboPlazo.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 500")
        C.oCboDestino.CargaTabla(, , "CredW_RecuperaCatalogoCodigos 2050")
        C.oDeudaFinan = RecuperaDeudasFinan_NewEval(pnCodEval, psNumDoc)
        C.oDeudaFinanCony = RecuperaDeudasFinanConyuge_NewEval(pnCodEval, psNumDoc)

        RecuperaEvalDeudasFinancieras_NewEval = C
        C = Nothing
    End Function

    Private Function RecuperaDeudasFinanConyuge_NewEval(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDeudasFinancConyuge_NewEval " & pnCodEval & ",'" & psNumDoc & "'")

        RecuperaDeudasFinanConyuge_NewEval = D
        D = Nothing
    End Function
    Private Function RecuperaDeudasFinan_NewEval(ByVal pnCodEval As Integer, ByVal psNumDoc As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaDeudasFinanc_NewFC " & pnCodEval & ",'" & psNumDoc & "'")

        RecuperaDeudasFinan_NewEval = D
        D = Nothing
    End Function

    Public Sub ActualizaDeudaFinanciera(ByVal pnCodEval As Integer, ByVal pnTipoPersonaAnt As Integer, ByVal pdFecVencAnt As Date, _
  ByVal pcFinancAnt As String, ByVal pnPrestamoAnt As Double, _
ByVal pnTipoPersona As Integer, ByVal pnSaldo As Double, ByVal pdFecVenc As Date, ByVal pcCuotaPag As String, _
ByVal pnCuotaMens As Double, ByVal pcFinanc As String, ByVal pcClasificacion As String, ByVal pnPrestamo As Double, _
ByVal pnDestino As Integer, ByVal pnPlazo As Integer, ByVal pcMoneda As String, ByVal pnCuotaFam As Double, _
ByVal pnCuotaNego As Double, ByVal pnCapAFecha As Double, ByVal pnTipoProd As Integer, ByVal pnCapFecLP As Double)

        Dim oCon As ComConecta.clsConecta

        oCon = New ComConecta.clsConecta
        Dim oComm As New SqlCommand("Pers_ActualizaDeudaFinanc", oCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@nTipoPersonaAn", SqlDbType.Int).Value = pnTipoPersonaAnt
        oComm.Parameters.Add("@dFecVencAn", SqlDbType.Date).Value = pdFecVencAnt
        oComm.Parameters.Add("@nPrestamoAn", SqlDbType.Money).Value = pnPrestamoAnt
        oComm.Parameters.Add("@cFinancAn", SqlDbType.VarChar, 100).Value = pcFinancAnt
        oComm.Parameters.Add("@nTipoPersona", SqlDbType.Int).Value = pnTipoPersona
        oComm.Parameters.Add("@nSaldo", SqlDbType.Money).Value = pnSaldo
        oComm.Parameters.Add("@dFecVenc", SqlDbType.Date).Value = pdFecVenc
        oComm.Parameters.Add("@cCuotaPag", SqlDbType.VarChar, 8).Value = pcCuotaPag
        oComm.Parameters.Add("@nCuotaMens", SqlDbType.Money).Value = pnCuotaMens
        oComm.Parameters.Add("@cFinanc", SqlDbType.VarChar, 100).Value = pcFinanc
        oComm.Parameters.Add("@cClasific", SqlDbType.VarChar, 30).Value = pcClasificacion
        oComm.Parameters.Add("@nPrestamos", SqlDbType.Money).Value = pnPrestamo
        oComm.Parameters.Add("@nDestino", SqlDbType.Int).Value = pnDestino
        oComm.Parameters.Add("@nPlazo", SqlDbType.Int).Value = pnPlazo
        oComm.Parameters.Add("@cMoneda", SqlDbType.VarChar, 20).Value = pcMoneda
        oComm.Parameters.Add("@nCuotaFam", SqlDbType.Money).Value = pnCuotaFam
        oComm.Parameters.Add("@nCuotaNeg", SqlDbType.Money).Value = pnCuotaNego
        oComm.Parameters.Add("@nCapAFec", SqlDbType.Money).Value = pnCapAFecha
        oComm.Parameters.Add("@nTipoProd", SqlDbType.Int).Value = pnTipoProd
        oComm.Parameters.Add("@nCapAFechaLP", SqlDbType.Money).Value = pnCapFecLP

        oComm.ExecuteNonQuery()
        oCon.CierraConexion()
        oComm = Nothing
        oCon = Nothing

    End Sub

End Class
