Imports System.Data
Imports System.Data.SqlClient
Public Class ClsAgencias
    Public Function ListarTodasAgencias() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_DevuelveTodasAgencias", , , False)
        ListarTodasAgencias = C
        C = Nothing
    End Function

    Public Sub InsertarAgencia(ByVal pncodage As Integer, ByVal psNomage As String, _
                               ByVal psDirecage As String, ByVal psTelefage As String, _
                               ByVal pscodagesbs As String, ByVal pncodplanilla As Integer, _
                               ByVal pnhabexpe As Integer, ByVal pncritica As Integer, _
                               ByVal pnhabdc As Integer, ByVal pccodzona As String, _
                               ByVal pccodzonasbs As String, ByVal pnNivelRiesgo As Integer, _
                               ByVal pnMaxEntNuevo As String, ByVal pnMaxEntRecurrente As Integer, _
                               ByVal pnHabFlujo As Integer, ByVal pnHabEval As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Pers_InsertaAgencia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.AddWithValue("@ncodage", pncodage)
        oComm.Parameters.AddWithValue("@cnomage", psNomage)
        oComm.Parameters.AddWithValue("@cdirecAge", psDirecage)
        oComm.Parameters.AddWithValue("@ctelefAge", psTelefage)
        oComm.Parameters.AddWithValue("@ccodagesbs", pscodagesbs)
        oComm.Parameters.AddWithValue("@ncodplanilla", pncodplanilla)
        oComm.Parameters.AddWithValue("@bhabilitaexperto", pnhabexpe)
        oComm.Parameters.AddWithValue("@bcritica", pncritica)
        oComm.Parameters.AddWithValue("@bhabdc", pnhabdc)
        oComm.Parameters.AddWithValue("@ccodzona", pccodzona)
        oComm.Parameters.AddWithValue("@ccodzonasbs", pccodzonasbs)
        oComm.Parameters.AddWithValue("@nNivelRiesgo", pnNivelRiesgo)
        oComm.Parameters.AddWithValue("@nMaxEntNuevo", pnMaxEntNuevo)
        oComm.Parameters.AddWithValue("@nMaxEntRecurrente", pnMaxEntRecurrente)
        oComm.Parameters.AddWithValue("@bhabNFC", pnHabFlujo)
        oComm.Parameters.AddWithValue("@bhabEval", pnHabEval)
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub ActualizarAgencia(ByVal pncodage As Integer, ByVal psNomage As String, _
                                 ByVal psDirecage As String, ByVal psTelefage As String, _
                                 ByVal pscodagesbs As String, ByVal pncodplanilla As Integer, _
                                 ByVal pnhabexpe As Integer, ByVal pncritica As Integer, _
                                 ByVal pnhabdc As Integer, ByVal pccodzona As String, _
                                ByVal pccodzonasbs As String, ByVal pnNivelRiesgo As Integer, _
                                 ByVal pnMaxEntNuevo As String, ByVal pnMaxEntRecurrente As Integer, _
                                 ByVal pnHabFlujo As Integer, ByVal pnHabEval As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Pers_Actualiza_Agencia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.AddWithValue("@ncodage", pncodage)
        oComm.Parameters.AddWithValue("@cnomage", psNomage)
        oComm.Parameters.AddWithValue("@cdirecAge", psDirecage)
        oComm.Parameters.AddWithValue("@ctelefAge", psTelefage)
        oComm.Parameters.AddWithValue("@ccodagesbs", pscodagesbs)
        oComm.Parameters.AddWithValue("@ncodplanilla", pncodplanilla)
        oComm.Parameters.AddWithValue("@bhabilitaexperto", pnhabexpe)
        oComm.Parameters.AddWithValue("@bcritica", pncritica)
        oComm.Parameters.AddWithValue("@bhabdc", pnhabdc)
        oComm.Parameters.AddWithValue("@ccodzona", pccodzona)
        oComm.Parameters.AddWithValue("@ccodzonasbs", pccodzonasbs)
        oComm.Parameters.AddWithValue("@nNivelRiesgo", pnNivelRiesgo)
        oComm.Parameters.AddWithValue("@nMaxEntNuevo", pnMaxEntNuevo)
        oComm.Parameters.AddWithValue("@nMaxEntRecurrente", pnMaxEntRecurrente)
        oComm.Parameters.AddWithValue("@bhabNFC", pnHabFlujo)
        oComm.Parameters.AddWithValue("@bhabEval", pnHabEval)
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub EliminarAgencia(ByVal codage As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("pers_eliminaragencia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@ncodage", SqlDbType.Int).Value = codage
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub
End Class
