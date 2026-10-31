Imports System.Data
Imports System.Data.SqlClient
Public Class clsEvalCreditoAgricola
    Public Function RecuperaControlesEvalCreditoAgricola() As ComTipos.clsEvalCreditoAgricola
        Dim C As ComTipos.clsEvalCreditoAgricola
        C = New ComTipos.clsEvalCreditoAgricola
        C.EvaluaCreditoAgricola.CargaTabla(, , "Eval_CredAgri_RecuperaCatalogoCogido 4042")
        RecuperaControlesEvalCreditoAgricola = C
        C = Nothing
    End Function
    Public Sub RegistraEvalCreditoAgricola(ByVal pnCodEval As Integer, ByVal pnTipoCul As Integer, _
    ByVal pnVariedad As String, ByVal pnExpAnios As Integer, ByVal pnFecCos As Date, _
    ByVal pnHectTot As Double, ByVal pnFecCom As Date, ByVal pnHectSemb As Double, _
     ByVal pnHectAlq As Double, ByVal pnHectBlanco As Double, ByVal pnProdHect As Double, _
     ByVal pnProdFan As Double, ByVal pnProdHCAnt1 As Double, ByVal pnProdHCAnt2 As Double, _
     ByVal pnProdFCAnt1 As Double, ByVal pnProdFCAnt2 As Double, ByVal pnInsumos As Double, ByVal pnManoObra As Double, ByVal pnMaqEqui As Double, ByVal pnAsiTec As Double, _
     Optional ByVal pnFinancEmpresa As Double = 0)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("Eval_CredAgri_Registra ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = pnCodEval
        oComm.Parameters.Add("@pnTipoCul", SqlDbType.Int).Value = pnTipoCul
        oComm.Parameters.Add("@pnVariedad", SqlDbType.NVarChar).Value = pnVariedad
        oComm.Parameters.Add("@pnExpAnios", SqlDbType.Int).Value = pnExpAnios
        oComm.Parameters.Add("@pnFecCos", SqlDbType.DateTime).Value = pnFecCos
        oComm.Parameters.Add("@pnHectTot", SqlDbType.Money).Value = pnHectTot
        oComm.Parameters.Add("@pnFecCom", SqlDbType.DateTime).Value = pnFecCom
        oComm.Parameters.Add("@pnHectSemb", SqlDbType.Money).Value = pnHectSemb
        oComm.Parameters.Add("@pnHectAlq", SqlDbType.Money).Value = pnHectAlq
        oComm.Parameters.Add("@pnHectBlanco", SqlDbType.Money).Value = pnHectBlanco
        oComm.Parameters.Add("@pnProdHect", SqlDbType.Money).Value = pnProdHect
        oComm.Parameters.Add("@pnProdFan", SqlDbType.Money).Value = pnProdFan
        oComm.Parameters.Add("@pnProdHCAnt1", SqlDbType.Money).Value = pnProdHCAnt1
        oComm.Parameters.Add("@pnProdHCAnt2", SqlDbType.Money).Value = pnProdHCAnt2
        oComm.Parameters.Add("@pnProdFCAnt1", SqlDbType.Money).Value = pnProdFCAnt1
        oComm.Parameters.Add("@pnProdFCAnt2", SqlDbType.Money).Value = pnProdFCAnt2
        oComm.Parameters.Add("@pnInsumos", SqlDbType.Money).Value = pnInsumos
        oComm.Parameters.Add("@pnManoObra", SqlDbType.Money).Value = pnManoObra
        oComm.Parameters.Add("@pnMaqEquipo", SqlDbType.Money).Value = pnMaqEqui
        oComm.Parameters.Add("@pnAsiTec", SqlDbType.Money).Value = pnAsiTec

        oComm.Parameters.Add("@nFinancEmpresa", SqlDbType.Money).Value = pnFinancEmpresa

        'oComm.Parameters.Add("@nBand", SqlDbType.Int)
        'oComm.Parameters.Item("@nBand").Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        'nBand = CInt(Trim(oComm.Parameters.Item(16).Value))

        oComm = Nothing
        objCon = Nothing
    End Sub

End Class
