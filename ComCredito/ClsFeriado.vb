Imports System.Data
Imports System.Data.SqlClient

Public Class ClsFeriado

    Public Function ESFeriado(ByVal pdFecha As Date) As Boolean
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        Dim nResul As Integer

        'Carga Nombres de Columnas de Tablas
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_Esferiado", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.VarChar).Value = pdFecha
        oComm.Parameters.Add("@nResul", SqlDbType.Bit)
        oComm.Parameters.Item("@nResul").Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()
        nResul = CInt(oComm.Parameters.Item("@NumRegs").Value.ToString)

    End Function
    Public Function ListarAgencias() As ComTipos.clsDataset
        Dim C As ComTipos.clsDataset
        C = New ComTipos.clsDataset
        C.CargaTabla(, , "Cred_ListarAgenciasF", , , False)
        ListarAgencias = C
        C = Nothing
    End Function

    Public Function ListarAgenciasFeriados(ByVal pdfecha As Date) As ComTipos.clsDataset
        Dim odatos As ComTipos.clsDataset
        odatos = New ComTipos.clsDataset
        odatos.CargaTabla(, , "Cred_Devuelve_AgenFeriados '" & Format(pdfecha, "dd/MM/yyyy") & "'", , , False)
        ListarAgenciasFeriados = odatos
        odatos = Nothing
    End Function

    Public Sub GrabarFeriadoAgencia(ByVal poLista As ComTipos.clsDataset, ByVal pdfecha As Date, ByVal psdescripcion As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_InsertaFeriadoAgencia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@xLista", SqlDbType.Xml).Value = poLista.DevuelveXML
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = pdfecha
        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar).Value = psdescripcion
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Sub EliminarFeriado(ByVal pdfecha As Date)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Cred_EliminaFeriado", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@dFecha", SqlDbType.Date).Value = pdfecha
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function Devolver_DescripFeriado(ByVal pdfecha As Date) As String
        Dim odatos As ComTipos.clsDataset
        odatos = New ComTipos.clsDataset
        odatos.CargaTabla(, , "Cred_Devuelve_DescripFeriado '" & Format(pdfecha, "dd/MM/yyyy") & "'", , , False)
        Devolver_DescripFeriado = odatos.ObtenerDato("cDescrip").ToString
        odatos = Nothing


    End Function
End Class
