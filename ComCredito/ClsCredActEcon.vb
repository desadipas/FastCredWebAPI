Option Explicit On

Public Class ClsCredActEcon

    Public Function RecuperaTiposAct() As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaTiposDEActEcon")
        RecuperaTiposAct = D
        D = Nothing

    End Function

    Public Function RecuperaSubTiposAct(ByVal psCodAct As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaSubTiposDEActEcon '" & psCodAct & "'")
        RecuperaSubTiposAct = D
        D = Nothing

    End Function

    Public Function RecuperaMantenimActEco(ByVal psSubTipoAct As String) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_RecuperaMantActividadesEcon '" & psSubTipoAct & "'")
        RecuperaMantenimActEco = D
        D = Nothing

    End Function


    Public Function RecuperaCorrelativoActEco(ByVal psSubTipoAct As String) As String
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand


        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_RecuperaCorrelativoActividadesEcon ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nSubTipoCodAct", SqlDbType.VarChar, 4).Value = psSubTipoAct
        oComm.Parameters.Add("@psCorrelativo", SqlDbType.VarChar, 20).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        RecuperaCorrelativoActEco = oComm.Parameters(1).Value

    End Function

    Public Sub RegistraActividadEconomica(ByVal psCodAct As String, ByVal psDescripcion As String)

    End Sub

End Class
