Imports System.Data
Imports System.Data.SqlClient

Public Class clsExperto

    Public Sub ExpertoActualizaParametros(ByVal oDataParam As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Tamiz_ActualizaParamentros ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@Data", SqlDbType.Xml).Value = oDataParam.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing

    End Sub

    Public Function ExpertoRecuperaParametros() As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "TamizRecuperaParametros ")
        ExpertoRecuperaParametros = D


        D = Nothing

    End Function
    Public Function RecuperaDatosRCC(ByVal psDoc As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "CredW_RecuperaRCC '" & psDoc & "'")
        RecuperaDatosRCC = D



        D = Nothing
    End Function

    Public Sub FiltrarGENESYSEXPERTO(ByVal psDocumento As String, _
                                        ByRef pnResultado As Integer, _
                                        ByRef psDescripcion As String, _
                                        ByVal pnCodAge As Integer, _
                                        ByVal psCodUsu As String, _
                                        ByVal psCliente As String, _
                                        ByVal psDocConyuge As String, _
                                        ByVal psConyuge As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_FiltraGENESYSEXPERTO ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@cDocumento", SqlDbType.VarChar, 20).Value = psDocumento
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@cCliente", SqlDbType.VarChar, 100).Value = psCliente
        oComm.Parameters.Add("@cDocConyuge", SqlDbType.VarChar, 20).Value = psDocConyuge
        oComm.Parameters.Add("@cConyuge", SqlDbType.VarChar, 20).Value = psConyuge
        oComm.ExecuteNonQuery()
        pnResultado = oComm.Parameters("@nResultado").Value
        psDescripcion = oComm.Parameters("@cDescripcion").Value

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub FiltrarGENESYSEXPERTO_Conyuge(ByVal psDocumento As String, _
                                    ByRef pnResultado As Integer, _
                                    ByRef psDescripcion As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_FiltraGENESYSEXPERTO_Conyuge", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cDocumento", SqlDbType.VarChar, 20).Value = psDocumento
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        pnResultado = oComm.Parameters("@nResultado").Value
        psDescripcion = oComm.Parameters("@cDescripcion").Value

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function ObtieneNombreCliente_Por_Documento(ByVal psDoc As String) As String
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Pers_ObtieneNombreCliente_Por_Documento '" & psDoc & "'")
        If D.nNumFilas > 0 Then
            ObtieneNombreCliente_Por_Documento = D.ObtenerDato("cCliente")
        Else
            ObtieneNombreCliente_Por_Documento = "NO EXISTE EL CLIENTE"
        End If
        D = Nothing
    End Function

    Public Function RecuperarAgenciasExperto() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_RecuperarAgenciasExperto")
        RecuperarAgenciasExperto = D

        D = Nothing
    End Function

    Public Sub ActualizarAgenciasExperto(ByVal poLista As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizarAgenciasExperto", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = poLista.DevuelveXML

        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub FiltraTamizCompraDeuda(ByVal psDocumento As String, _
                                ByVal pnCodAge As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_FiltraTamizCompraDeuda ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@cDocumento", SqlDbType.VarChar, 20).Value = psDocumento
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub FiltrarClienteDATACREDITO(ByVal oDatos As ComTipos.clsDataset, _
                                      ByVal sDocumento As String, _
                         ByVal sConsulta As String, _
                         ByRef nCodFiltro As Integer, _
                         Optional ByVal nMontoFiltro As Double = 0.0, _
                         Optional ByVal sDocConyuge As String = "", _
                         Optional ByVal nCodAge As Integer = 0, _
                         Optional ByVal sUsuario As String = "", _
                         Optional ByVal nRelacion As Integer = 0, _
                         Optional ByVal nCodFiltroTAMIZ As Integer = 0)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_FiltrarClienteDATACREDITO ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        'oComm.CommandTimeout = 0
        oComm.Parameters.Add("@oDatos", SqlDbType.Xml).Value = oDatos.DevuelveXML
        oComm.Parameters.Add("@cDocumento", SqlDbType.VarChar, 20).Value = sDocumento
        oComm.Parameters.Add("@cConsulta", SqlDbType.Text).Value = sConsulta
        oComm.Parameters.Add("@nCodFiltro", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@cDocConyuge", SqlDbType.VarChar, 20).Value = sDocConyuge 'ARCV 22-02-2012
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge 'ARCV 28-02-2012
        oComm.Parameters.Add("@cUsuario", SqlDbType.VarChar, 20).Value = sUsuario 'ARCV 28-02-2012
        oComm.Parameters.Add("@nMontoFiltro", SqlDbType.Money).Value = nMontoFiltro  'CPUA 25-04-2012
        oComm.Parameters.Add("@nRelacion", SqlDbType.Int).Value = nRelacion 'ARCV 29-08-2012
        oComm.Parameters.Add("@nCodFiltroTAMIZ", SqlDbType.Int).Value = nCodFiltroTAMIZ 'ARCV 15-10-2012

        oComm.ExecuteNonQuery()
        nCodFiltro = oComm.Parameters("@nCodFiltro").Value


        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function BuscarFiltroDATACREDITO(ByVal nCodFiltro As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarFiltroDATACREDITO " & nCodFiltro)
        BuscarFiltroDATACREDITO = D

        D = Nothing
    End Function

    Public Function BuscarFiltroDATACREDITO_Usuario(ByVal nCodFiltro As Integer, _
                                ByVal psUsuario As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarFiltroDATACREDITO_Usuario " & nCodFiltro & ",'" & psUsuario & "'")
        BuscarFiltroDATACREDITO_Usuario = D

        D = Nothing
    End Function

    Public Function BuscarExcepcionDATACREDITO(ByVal nCodFiltro As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarExcepcionDATACREDITO " & nCodFiltro)
        BuscarExcepcionDATACREDITO = D

        D = Nothing
    End Function

    Public Sub RegistraDataCreditoBloqueo(ByVal psDocumento As String, _
                            ByVal pdFecha As Date, _
                            ByVal pnTipoBloqueo As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("cred_RegistraDataCreditoBloqueo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@cDocumento", SqlDbType.VarChar, 20).Value = psDocumento
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = pdFecha
        oComm.Parameters.Add("@nTipoBloqueo", SqlDbType.Int).Value = pnTipoBloqueo
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function BuscarFiltroDATACREDITO_Documento(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarFiltroDATACREDITO_Documento '" & psDocumento & "'")
        BuscarFiltroDATACREDITO_Documento = D
        D = Nothing
    End Function

    Public Function BuscarFiltroDATACREDITO_Dia(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarFiltroDATACREDITO_Dia '" & psDocumento & "'")
        BuscarFiltroDATACREDITO_Dia = D
        D = Nothing
    End Function

    Public Sub RegistraDataCreditoRevision(ByVal pnCodFiltro As Integer, _
                        ByVal pdFecha As Date, _
                        ByVal psUsuario As String, _
                    ByVal psComentario As String, _
                    ByVal pbAprobado As Boolean)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RegistraDataCreditoRevision ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@nCodFiltro", SqlDbType.Int).Value = pnCodFiltro '
        oComm.Parameters.Add("@cUsuarioRev", SqlDbType.VarChar, 20).Value = psUsuario
        oComm.Parameters.Add("@dFechaRev", SqlDbType.DateTime).Value = pdFecha
        oComm.Parameters.Add("@cComentarioRev", SqlDbType.VarChar, 200).Value = psComentario
        oComm.Parameters.Add("@bAprobadoRev", SqlDbType.Bit).Value = IIf(pbAprobado, 1, 0)
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function ListarDataCreditoFiltro(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_ListarDataCreditoFiltro '" & psDocumento & "'")
        ListarDataCreditoFiltro = D
        D = Nothing
    End Function

    Public Function BuscarListaCredDataCreditoBloqueo(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarListaCredDataCreditoBloqueo '" & psDocumento & "'")
        BuscarListaCredDataCreditoBloqueo = D
        D = Nothing
    End Function

    Public Sub ActualizarCredDataCreditoBloqueo(ByVal pnCodFiltro As Integer, _
                    ByVal pdFecha As Date, _
                    ByVal psUsuario As String, _
                ByVal psComentario As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ActualizarCredDataCreditoBloqueo ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@nCodigo", SqlDbType.Int).Value = pnCodFiltro '
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = psUsuario
        oComm.Parameters.Add("@dFecReg", SqlDbType.DateTime).Value = pdFecha
        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 200).Value = psComentario
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function BuscarFiltroDATACREDITO_Historico(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarFiltroDATACREDITO_Historico '" & psDocumento & "'")
        BuscarFiltroDATACREDITO_Historico = D
        D = Nothing
    End Function

    Public Function BuscarDataCreditoRevision(ByVal psDocumento As String) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarDataCreditoRevision '" & psDocumento & "'")
        BuscarDataCreditoRevision = D
        D = Nothing
    End Function

    Public Sub ModificaDataCreditoDecision(ByVal pnCodFiltro As Integer, _
                   ByVal psClasificacion As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_ModificaDataCreditoDecision ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@nCodFiltro", SqlDbType.Int).Value = pnCodFiltro '
        oComm.Parameters.Add("@cClasificacion", SqlDbType.Char, 1).Value = psClasificacion
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function DevuelveSustentoDataCreditoxFiltro(ByVal psDocumento As String, ByVal pnCodFiltro As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "DataCreditoDevuelveSustento '" & psDocumento & "', " & pnCodFiltro)
        DevuelveSustentoDataCreditoxFiltro = D
        D = Nothing
    End Function

    Public Sub DataCreditoRegistraSustento(ByVal psDocumento As String, ByVal pnCodFiltro As Integer, _
               ByVal psInfoAlertas As String, ByVal psMitiAlertas As String, ByVal psSustentoDoc As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("DataCreditoRegistraSustento ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@sDocumento", SqlDbType.VarChar, 13).Value = psDocumento
        oComm.Parameters.Add("@nCodFiltro", SqlDbType.Int).Value = pnCodFiltro
        oComm.Parameters.Add("@sInfoAlertas", SqlDbType.VarChar, 1000).Value = psInfoAlertas
        oComm.Parameters.Add("@sMitiAlertas", SqlDbType.VarChar, 1000).Value = psMitiAlertas
        oComm.Parameters.Add("@sSustentoDoc", SqlDbType.VarChar, 1000).Value = psSustentoDoc
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function DevuelveRevisionDataCreditoxFiltro(ByVal pnCodFiltro As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "DataCreditoDevuelveComentariosRevision " & pnCodFiltro)
        DevuelveRevisionDataCreditoxFiltro = D
        D = Nothing
    End Function

    Public Function DataCreditoDevuelveResupesta(ByVal pnCodRespuesta As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "DataCreditoDevuelveResupesta " & pnCodRespuesta)
        DataCreditoDevuelveResupesta = D
        D = Nothing
    End Function

    Public Sub FiltroTAMIZGENESYS(ByVal psDocumento As String, _
                                    ByRef pnResultado As Integer, _
                                    ByRef psDescripcion As String, _
                                    ByVal pnCodAge As Integer, _
                                    ByVal psCodUsu As String, _
                                    ByVal psCliente As String, _
                                    ByVal psDocConyuge As String, _
                                    ByVal psConyuge As String, _
                                    ByRef pnNroOpe As Integer, _
                                    ByVal pnRelacion As Integer, _
                                    ByVal pnMontoFiltro As Double, _
                                    ByRef pnCodFiltroDC As Integer, _
                                    ByVal pnCodPers As Integer, _
                                    ByVal pnCodCred As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_FiltroTAMIZGENESYS ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@cDocumento", SqlDbType.VarChar, 20).Value = psDocumento
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = psCodUsu
        oComm.Parameters.Add("@cCliente", SqlDbType.VarChar, 100).Value = psCliente
        oComm.Parameters.Add("@cDocConyuge", SqlDbType.VarChar, 20).Value = psDocConyuge
        oComm.Parameters.Add("@cConyuge", SqlDbType.VarChar, 20).Value = psConyuge
        oComm.Parameters.Add("@nNroOpe", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nRelacion", SqlDbType.Int).Value = pnRelacion
        oComm.Parameters.Add("@nMontoFiltro", SqlDbType.Money).Value = pnMontoFiltro
        oComm.Parameters.Add("@nCodFiltroDC", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = pnCodPers
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred

        oComm.ExecuteNonQuery()
        pnResultado = oComm.Parameters("@nResultado").Value
        psDescripcion = oComm.Parameters("@cDescripcion").Value
        pnNroOpe = oComm.Parameters("@nNroOpe").Value
        pnCodFiltroDC = oComm.Parameters("@nCodFiltroDC").Value

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function BuscarFiltroDATACREDITO_Revisado(ByVal psDocumento As String, _
                                                    ByVal pnCodCred As Integer, _
                                                    ByVal pnCodAge As Integer) As ComTipos.clsDataset

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset

        D.CargaTabla(, , "Cred_BuscarFiltroDATACREDITO_Revisado '" & psDocumento & "', " & pnCodCred & "," & pnCodAge)
        BuscarFiltroDATACREDITO_Revisado = D
        D = Nothing
    End Function

    Public Function RecuperaRelacionTopeFiltro() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset

        D = New ComTipos.clsDataset

        D.CargaTabla(, , "cred_RecuperaRelacionTopeFiltro ")
        RecuperaRelacionTopeFiltro = D
        D = Nothing
    End Function

    Public Sub GrabarRelacionTopeFiltro(ByVal oDatos As ComTipos.clsDataset)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("cred_GrabarRelacionTopeFiltro ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0
        oComm.Parameters.Add("@oDatos", SqlDbType.Xml).Value = oDatos.DevuelveXML
        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub
End Class
