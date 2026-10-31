Public Class ClsPlanilla

    Public Sub EliminaCreditoDERecaudador(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, _
        ByVal pnCodPers As Integer)

        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_EliminaRecaudadorDECredito " & pnCodAge & "," & pnCodCred & "," & pnCodPers)
        D = Nothing


    End Sub


    Public Sub RegistraRecaudadorACredito(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnCodPers As Integer)
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RegistraRecaudadorACredito " & pnCodAge & "," & pnCodCred & "," & pnCodPers)
        D = Nothing

    End Sub

    Public Function RecuperaCreditosDERecaudador(ByVal pnCodpers As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaCreditosDeRecaudador " & pnCodpers)
        RecuperaCreditosDERecaudador = D
        D = Nothing

    End Function


    Public Function CargaRecaudador(ByVal pnCodAge As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaRecaudador " & pnCodAge)
        CargaRecaudador = D
        D = Nothing

    End Function

    Public Function CargarCombosPlanilla(ByVal pnCodAge As Integer) As ComTipos.ClsCredPlanilla
        Dim C As ComTipos.ClsCredPlanilla
        C = New ComTipos.ClsCredPlanilla
        C.cboAgencia.CargaTabla(, , "CredW_Recupera_Agencias")
        C.cboProducto.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '1500' ")
        C.cboSubProd.CargaTabla(, , "CredW_RecuperaCatalogoCodigos '1600' ")
        C.cboAnalista.CargaTabla(, , "Cred_RecuperaAnalista " & pnCodAge)
        C.cboRecaudador.CargaTabla(, , "Cred_RecuperaRecaudador " & pnCodAge)
        C.cboRuta.CargaTabla(, , "Cred_RecuperaRuta")
        CargarCombosPlanilla = C
        C = Nothing
    End Function


    Public Function CARGARDETALLESPLANILLA(ByVal CODE As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperarPlanilla " & CODE)
        Return D
        D = Nothing
    End Function


    Public Function CARGARDETALLESPLANILLAESTADO(ByVal CODE As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperarPlanillaEstado " & CODE)
        'MsgBox(D.DevuelveXML)
        Return D
        D = Nothing
    End Function
    Public Function GRABARDETALLESPLANILLA(ByVal ITF As Double, ByVal IGV As Double, ByVal ACTUALIZA As String, ByVal CODRUTA As Integer, ByVal CODEAGE As Integer, ByVal CODEANALISTA As Integer, ByVal CODEPROD As Integer, ByVal CODESUBPROD As Integer, ByVal FECHA As Date, ByVal pnCodRecaudador As Integer) As Integer
        Dim C As New ComConecta.clsConecta
        'Dim D
        'Dim D As New ComTipos.clsDataset
        Dim oCmd As New SqlClient.SqlCommand
        oCmd.CommandText = "Cred_RegistrarPlanilla"
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Connection = C.DameConexion
        oCmd.Parameters.Add("@Actualiza", SqlDbType.VarChar).Value = ACTUALIZA
        oCmd.Parameters.Add("@pdFecha", SqlDbType.DateTime).Value = FECHA.ToShortDateString
        oCmd.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = CODEAGE
        'oCmd.Parameters.Add("@pnCodRecaudador", SqlDbType.Int).Value = CODERECAUDADOR
        oCmd.Parameters.Add("@pnRuta", SqlDbType.Int).Value = CODRUTA
        oCmd.Parameters.Add("@pnCodAnalista", SqlDbType.Int).Value = CODEANALISTA
        oCmd.Parameters.Add("@pnCodProd", SqlDbType.Int).Value = CODEPROD
        oCmd.Parameters.Add("@pnCodSubProd", SqlDbType.Int).Value = CODESUBPROD
        oCmd.Parameters.Add("@pnValorItf", SqlDbType.Money).Value = ITF
        oCmd.Parameters.Add("@pnValorIGV", SqlDbType.Money).Value = IGV
        oCmd.Parameters.Add("@CODE", SqlDbType.Int).Direction = ParameterDirection.Output
        oCmd.Parameters.Add("@pnCodRecaudador", SqlDbType.Int).Value = pnCodRecaudador

        oCmd.ExecuteNonQuery()
        'D.CargaTabla(, , "Cred_RecuperarPlanilla " & oCmd.Parameters.Item("@CODE").Value)
        Return Val(oCmd.Parameters.Item("@CODE").Value)
        'D = Nothing
        C.CierraConexion()
        C = Nothing
        oCmd = Nothing
    End Function

    Public Function ACTUALIZARDETALLESPLANILLA(ByVal CODEPLA As Integer, ByVal oXML As ComTipos.clsDataset, ByVal o2XML As ComTipos.clsDataset, _
                                            ByVal sUSER As String) As ComTipos.clsDataset
        'Public Function ACTUALIZARDETALLESPLANILLA(ByVal oXML As ComTipos.clsDataset) As String
        Dim C As New ComConecta.clsConecta
        Dim D As New ComTipos.clsDataset
        Dim oCmd As SqlClient.SqlCommand
        oCmd = New SqlClient.SqlCommand("Cred_ActualizarCrediDiario ", C.DameConexion)
        'oCmd.CommandText = "Cred_ActualizarCrediDiario"
        'oCmd.CommandText = "Cred_PRUEBAXML"
        oCmd.CommandType = CommandType.StoredProcedure
        'oCmd.Connection = C.DameConexion
        oXML.AgregaColumna("Rows")
        For i As Integer = 0 To oXML.nNumFilas - 1
            oXML.AsignarDato(i.ToString, "Rows", i)
        Next
        o2XML.AgregaColumna("Row")
        For j As Integer = 0 To o2XML.nNumFilas - 1
            o2XML.AsignarDato(j.ToString, "Row", j)
        Next
        'MsgBox(oXML.DevuelveXML)
        'oCmd.Parameters.Add("@CODPLA", SqlDbType.Int).Value = CODEPLA
        oCmd.Parameters.Add("@Datos", SqlDbType.Xml).Value = oXML.DevuelveXML
        oCmd.Parameters.Add("@MAXROWS", SqlDbType.Int).Value = oXML.nNumFilas
        'MsgBox(oCmd.Parameters.Item("@Datos").Value)
        oCmd.Parameters.Add("@cERROR", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

        oCmd.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUSER 'ARCV 17-11-2009
        'Try
        oCmd.ExecuteNonQuery()
        'Catch ex As Exception
        ' MsgBox(ex.Message)
        'End Try

        'MsgBox("Se mostrarán los Créditos que se ha podido Pagar Satisfactoriamente" & Chr(13) & oCmd.Parameters.Item("@cERROR").Value.ToString)

        'Dim o2xml As ComTipos.clsDataset
        D.CargaTabla(, , "Cred_MostrarCreditosPagado " & CODEPLA & ",'" & o2XML.DevuelveXML & "'," & o2XML.nNumFilas)
        Return D
        'Return oCmd.Parameters.Item("@cERROR").Value.ToString
        C.CierraConexion()
        C = Nothing
        oCmd = Nothing

        'Dim D As ComTipos.clsDataset
        'D = New ComTipos.clsDataset
        'D.CargaTabla(, , "Cred_ActualizarCrediDiario " & CODEPLA & ",'" & oXML.DevuelveXML & "'," & oXML.nNumFilas)
        'Return D
        'D = Nothing




        ''INICIO DE NUEVA FUNCION
        'Dim ContFilas As Integer
        'Dim ContColum As Integer
        'Dim Colu() As String
        'Dim Dat(,) As String
        'Dim nCur As Integer
        'Dim C As New ComConecta.clsConecta
        'Dim oCmd As New SqlClient.SqlCommand
        'oCmd.Connection = C.DameConexion
        'oCmd.CommandType = CommandType.StoredProcedure
        'oCmd.CommandText = "Cred_ActualizarCrediDiario"
        'oCmd.Parameters.Add("@CODPLA", SqlDbType.Int).Value = CODEPLA
        'oCmd.Parameters.Add("@Datos", SqlDbType.Xml).Value = oXML.DevuelveXML
        'oCmd.Parameters.Add("@MAXROWS", SqlDbType.Int).Value = oXML.nNumFilas
        ''oCmd.Parameters.Item("@NumRegs").Direction = ParameterDirection.Output
        ''MsgBox(oComm.Parameters.Item("@sCadena").Value)
        ''oCmd.ExecuteNonQuery()
        ''ContFilas = CInt(oCmd.Parameters.Item("@NumRegs").Value.ToString)
        ''dr = oCmd.ExecuteReader(CommandBehavior.SingleResult)
        'Dim rdr As SqlClient.SqlDataReader = oCmd.ExecuteReader(CommandBehavior.Default)
        ''Carga Columnas

        '' While rdr.Read()
        ''int contactID = rdr.GetInt32(rdr.GetOrdinal("ContactID"));
        '' string firstName = rdr.GetString(rdr.GetOrdinal("FirstName"));
        '' Response.Write(firstName + " (" + contactID.ToString() + ")");
        '' End While
        '' rdr.Close()



        ''ContColum = rdr.FieldCount
        ''ReDim Preserve Colu(ContColum)
        ''For i As Integer = 1 To ContColum
        '' Colu(i) = Trim(rdr.GetName(i - 1))
        ''Next
        ' ''Carga Datos
        ''ReDim Dat(ContFilas, ContColum)
        ''nCur = 0
        ''While rdr.Read
        '' nCur = nCur + 1
        '' For i As Integer = 1 To ContColum
        '' Dat(nCur, i) = Trim(IIf(IsDBNull(rdr.Item(Colu(i))), "", rdr.Item(Colu(i))))
        '' Next i
        ''End While
        ''ContFilas = nCur
        'Return rdr
        'rdr.Close()
        'rdr = Nothing
        'oCmd.Dispose()
        'oCmd = Nothing


        ''FIN DE FUNCION






    End Function
    Public Function BUSCARPLANILLAS(ByVal FECDESDE As DateTime, ByVal FECHASTA As DateTime) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_BuscarPlanillas '" & FECDESDE.ToShortDateString & "','" & FECHASTA.ToShortDateString & "'")
        Return D
        D = Nothing
    End Function

    Public Function BUSCARRUTAS() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_BuscarRutas")
        Return D
        D = Nothing
    End Function
    Public Sub REGISTRARRUTAS(ByVal CODREC As Integer, ByVal NOMRUTA As String, ByVal COMENTARIO As String)
        Dim C As New ComConecta.clsConecta
        Dim oCmd As New SqlClient.SqlCommand
        oCmd.CommandText = "Cred_RegistrarRuta"
        oCmd.Connection = C.DameConexion
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add("@CodRec", SqlDbType.Int).Value = CODREC
        oCmd.Parameters.Add("@NomRut", SqlDbType.VarChar).Value = NOMRUTA
        oCmd.Parameters.Add("@Comentario", SqlDbType.VarChar).Value = COMENTARIO
        oCmd.ExecuteNonQuery()
        oCmd = Nothing
        C = Nothing
    End Sub

    Public Sub ACTUALIZARRUTAS(ByVal CODRUTA As Integer, ByVal CODREC As Integer, ByVal NOMRUTA As String, ByVal COMENTARIO As String)
        Dim C As New ComConecta.clsConecta
        Dim oCmd As New SqlClient.SqlCommand
        oCmd.CommandText = "Cred_ActualizarRuta"
        oCmd.Connection = C.DameConexion
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add("@CodRuta", SqlDbType.Int).Value = CODRUTA
        oCmd.Parameters.Add("@CodRec", SqlDbType.Int).Value = CODREC
        oCmd.Parameters.Add("@NomRut", SqlDbType.VarChar).Value = NOMRUTA
        oCmd.Parameters.Add("@Comentario", SqlDbType.VarChar).Value = COMENTARIO
        oCmd.ExecuteNonQuery()
        oCmd = Nothing
        C = Nothing
    End Sub

    Public Function BUSCARZONAS(ByVal CODERUTA As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_BuscarZonas " & CODERUTA)
        Return D
        D = Nothing
    End Function

    Public Sub REGISTRARZONAS(ByVal CODRUTA As Integer, ByVal CODZONA As String)
        Dim C As New ComConecta.clsConecta
        Dim oCmd As New SqlClient.SqlCommand
        oCmd.CommandText = "Cred_RegistrarZonaRuta"
        oCmd.Connection = C.DameConexion
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add("@CodRut", SqlDbType.Int).Value = CODRUTA
        oCmd.Parameters.Add("@CodZona", SqlDbType.VarChar).Value = CODZONA
        'oCmd.Parameters.Add("@Error", SqlDbType.VarChar).Direction = ParameterDirection.Output
        oCmd.ExecuteNonQuery()
        oCmd = Nothing
        C = Nothing
    End Sub
    Public Sub ELIMINARRUTA(ByVal CODRUTA As Integer)
        Dim C As New ComConecta.clsConecta
        Dim oCmd As New SqlClient.SqlCommand
        oCmd.CommandText = "Cred_EliminarRuta"
        oCmd.Connection = C.DameConexion
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add("@CodRuta", SqlDbType.Int).Value = CODRUTA
        oCmd.ExecuteNonQuery()
        oCmd = Nothing
        C = Nothing
    End Sub
    Public Sub ELIMINARZONA(ByVal CODRUTA As Integer, ByVal CODZONA As String)
        Dim C As New ComConecta.clsConecta
        Dim oCmd As New SqlClient.SqlCommand
        oCmd.CommandText = "Cred_EliminarZonaRuta"
        oCmd.Connection = C.DameConexion
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add("@CodRuta", SqlDbType.Int).Value = CODRUTA
        oCmd.Parameters.Add("@CodZona", SqlDbType.VarChar).Value = CODZONA
        'MsgBox(oCmd.Parameters.Item("@CodRuta").Value)
        oCmd.ExecuteNonQuery()
        oCmd = Nothing
        C = Nothing
    End Sub
    Public Function CARGARCLIENTERUTA(ByVal CODE As Integer) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaHojaRuta " & Val(CODE))
        Return D
        D = Nothing
    End Function
    Public Sub ACTUALIZARRUTAPORCLIENTE(ByVal CODE As Integer, ByVal oXML As ComTipos.clsDataset)
        Dim C As New ComConecta.clsConecta
        Dim oCmd As New SqlClient.SqlCommand
        oXML.AgregaColumna("Rows")
        For i As Integer = 0 To oXML.nNumFilas - 1
            oXML.AsignarDato(i.ToString, "Rows", i)
        Next
        oCmd.CommandText = "Cred_ActualizarRutaCliente"
        oCmd.Connection = C.DameConexion
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add("@CodRuta", SqlDbType.Int).Value = CODE
        oCmd.Parameters.Add("@Datos", SqlDbType.Xml).Value = oXML.DevuelveXML
        oCmd.Parameters.Add("@MAXROWS", SqlDbType.Int).Value = oXML.nNumFilas - 1
        'MsgBox(oCmd.Parameters.Item("@CodRuta").Value)
        oCmd.ExecuteNonQuery()
        oCmd = Nothing
        C = Nothing
    End Sub
    'Public Function OBTENERDETPLANILLA(ByVal CODE As Integer) As DataTable
    ' Dim C As New ComConecta.clsConecta
    ' Dim oDA As SqlClient.SqlDataAdapter
    ' Dim DT As New DataTable
    ' Dim oCmd As New SqlClient.SqlCommand
    ' oCmd.CommandText = "Cred_RecuperarPlanillaEstado"
    ' oCmd.CommandType = CommandType.StoredProcedure
    ' oCmd.Connection = C.DameConexion
    ' oCmd.Parameters.Add("@CODE", SqlDbType.Int).Value = CODE
    ' oDA = New SqlClient.SqlDataAdapter(oCmd.ExecuteNonQuery, C.DameConexion)
    ' oDA.Fill(DT)
    ' Return DT
    'End Function


End Class
