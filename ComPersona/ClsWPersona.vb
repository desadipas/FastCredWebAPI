Public Class ClsWPersona

    Public P As ComTipos.ClsWCargaPersona

    Public Function RecuperaSoloDatos(ByVal psDoc As String) As ComTipos.ClsWCargaPersona
        P = New ComTipos.ClsWCargaPersona
        P.Datos.CargaTabla(, , "CredWCargaDatosPersona '" + psDoc + "'")
        RecuperaSoloDatos = P
    End Function

    Public Function RecuperaDatosPersona(ByVal psDoc As String) As ComTipos.ClsWCargaPersona
        P = New ComTipos.ClsWCargaPersona
        P.Datos.CargaTabla(, , "CredWCargaDatosPersona '" + psDoc + "'")
        P.cboCIIU.CargaTabla(, , "CredW_RecuperaCIIU")
        P.cboEstadoCiv.CargaTabla(, , "Cred_CatalogoCod '100'")
        P.cboDpto.CargaTabla("Zonas", "cCodZona like '__0000000000' order by cnomzona")
        P.cboSexo.CargaTabla(, , "Cred_CatalogoCod '2030'")
        P.cboTipDocs.CargaTabla(, , "Cred_CatalogoCod '1400'")
        P.cboTipoPersona.CargaTabla(, , "Cred_CatalogoCod '1800'")
        RecuperaDatosPersona = P
    End Function

    Public Function RecuperaBusquedaPersona(ByVal psApeNom As String) As ComTipos.clsDataset
        Dim O As ComTipos.clsDataset

        O = New ComTipos.clsDataset
        O.CargaTabla(, , "CredWBuscaPersona '" & psApeNom & "%'")
        RecuperaBusquedaPersona = O

    End Function

    Public Sub GrabaDatos(ByVal pDatos As ComTipos.ClsWCargaPersona, ByVal pnTipoNewUpd As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("CredW_GuardaPersona", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cApePat", SqlDbType.VarChar)
        oComm.Parameters("@cApePat").Value = pDatos.Datos.ObtenerDato("cApePat")

        oComm.Parameters.Add("@cApeMat", SqlDbType.VarChar)
        oComm.Parameters("@cApeMat").Value = pDatos.Datos.ObtenerDato("cApeMat")

        oComm.Parameters.Add("@cNombres", SqlDbType.VarChar)
        oComm.Parameters("@cNombres").Value = pDatos.Datos.ObtenerDato("cNombres")

        oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar)
        oComm.Parameters("@cTelefono").Value = pDatos.Datos.ObtenerDato("cTelefono")

        oComm.Parameters.Add("@nSexo", SqlDbType.VarChar)
        oComm.Parameters("@nSexo").Value = pDatos.Datos.ObtenerDato("nSexo")

        oComm.Parameters.Add("@cDNI", SqlDbType.VarChar)
        oComm.Parameters("@cDNI").Value = IIf(pDatos.Datos.ObtenerDato("cDNI") = Nothing, "", pDatos.Datos.ObtenerDato("cDNI"))

        oComm.Parameters.Add("@cRUC", SqlDbType.VarChar)
        oComm.Parameters("@cRUC").Value = IIf(pDatos.Datos.ObtenerDato("cRUC") = Nothing, "", pDatos.Datos.ObtenerDato("cRUC"))

        oComm.Parameters.Add("@cDirReferencia", SqlDbType.VarChar)
        oComm.Parameters("@cDirReferencia").Value = pDatos.Datos.ObtenerDato("cDirReferencia")

        oComm.Parameters.Add("@cZona", SqlDbType.VarChar)
        oComm.Parameters("@cZona").Value = pDatos.Datos.ObtenerDato("cZona")

        oComm.Parameters.Add("@cCIIU", SqlDbType.VarChar)
        oComm.Parameters("@cCIIU").Value = pDatos.Datos.ObtenerDato("cCIIU")

        oComm.Parameters.Add("@dFecNac", SqlDbType.VarChar)
        oComm.Parameters("@dFecNac").Value = pDatos.Datos.ObtenerDato("dFecNac")

        oComm.Parameters.Add("@nTipoPersona", SqlDbType.VarChar)
        oComm.Parameters("@nTipoPersona").Value = pDatos.Datos.ObtenerDato("nTipoPersona")

        oComm.Parameters.Add("@nEstadoCivil", SqlDbType.VarChar)
        oComm.Parameters("@nEstadoCivil").Value = pDatos.Datos.ObtenerDato("nEstadoCivil")

        oComm.Parameters.Add("@nCodPers", SqlDbType.VarChar)
        oComm.Parameters("@nCodPers").Value = pDatos.Datos.ObtenerDato("nCodPers")

        oComm.Parameters.Add("@nCodAge", SqlDbType.VarChar)
        oComm.Parameters("@nCodAge").Value = pDatos.Datos.ObtenerDato("nCodAge")

        oComm.Parameters.Add("@nTipoAct", SqlDbType.Int)
        oComm.Parameters("@nTipoAct").Value = pnTipoNewUpd

        oComm.Parameters.Add("@cRazonSocNeg", SqlDbType.VarChar)
        oComm.Parameters("@cRazonSocNeg").Value = pDatos.Datos.ObtenerDato("cRazonSocNeg")

        oComm.Parameters.Add("@cDireccNeg", SqlDbType.VarChar)
        oComm.Parameters("@cDireccNeg").Value = pDatos.Datos.ObtenerDato("cDireccNeg")

        oComm.Parameters.Add("@cTelefonoNeg", SqlDbType.VarChar)
        oComm.Parameters("@cTelefonoNeg").Value = pDatos.Datos.ObtenerDato("cTelefonoNeg")

        oComm.Parameters.Add("@nCodPersRela", SqlDbType.Int)
        oComm.Parameters("@nCodPersRela").Value = CInt(pDatos.Datos.ObtenerDato("nCodPersRel"))

        oComm.ExecuteNonQuery()

    End Sub

End Class
